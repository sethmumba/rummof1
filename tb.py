"""
Human-like typing bot (step 1 + 2: typing engine and recorder).
Windows + VS Code. Every logged event carries synthetic=True.

Install:  pip install pynput
Run:      python typing_bot.py template.py
Then click into an empty VS Code editor during the countdown.
Press F12 at any time to abort.
"""
import json
import math
import random
import sys
import time
from dataclasses import dataclass

from pynput import keyboard

kb = keyboard.Controller()
abort = False


@dataclass
class Persona:
    name: str = "average"
    wpm: float = 55          # base typing speed
    sigma: float = 0.35      # timing variability (log-normal spread)
    typo_rate: float = 0.02  # chance per letter
    line_pause: tuple = (0.4, 2.0)   # thinking pause after a newline
    word_pause_prob: float = 0.06    # hesitation at word boundaries
    drift: float = 0.08      # slow speed drift over the session


PERSONAS = {
    "average": Persona(),
    "fast": Persona("fast", wpm=90, sigma=0.25, typo_rate=0.01),
    "hunt_and_peck": Persona("hunt_and_peck", wpm=28, sigma=0.5, typo_rate=0.04),
    "corrector": Persona("corrector", wpm=60, sigma=0.3, typo_rate=0.06),
}

FAST_PAIRS = {"th", "he", "in", "er", "an", "re", "on", "at", "en", "nd", "es", "or"}
SLOW_CHARS = set("{}()[]<>_|\\\"'~^&*#@$%:+=")
NEIGHBORS = {
    "q": "wa", "w": "qes", "e": "wrd", "r": "etf", "t": "ryg", "y": "tuh",
    "u": "yij", "i": "uok", "o": "ipl", "p": "o", "a": "qsz", "s": "awdx",
    "d": "sefc", "f": "drgv", "g": "fthb", "h": "gyjn", "j": "hukm",
    "k": "jil", "l": "ko", "z": "ax", "x": "zsc", "c": "xdv", "v": "cfb",
    "b": "vgn", "n": "bhm", "m": "nj",
}

t0 = time.perf_counter()
log_file = None


def log(source, event, key, **extra):
    rec = {"t": round(time.perf_counter() - t0, 4), "source": source,
           "event": event, "key": key, "synthetic": True, **extra}
    log_file.write(json.dumps(rec) + "\n")


def char_delay(persona, prev, ch, elapsed):
    mean = 60.0 / (persona.wpm * 5)
    mean *= 1 + persona.drift * math.sin(elapsed / 90.0)   # slow drift
    if prev and (prev + ch).lower() in FAST_PAIRS:
        mean *= 0.75
    if ch in SLOW_CHARS or ch.isupper():
        mean *= 1.4
    if prev == " ":
        mean *= 1.15
    mu = math.log(mean) - persona.sigma ** 2 / 2
    return random.lognormvariate(mu, persona.sigma)


def press(key, hold=None):
    """Press and release with a realistic key-hold time; log both edges."""
    name = key if isinstance(key, str) else str(key)
    hold = hold if hold is not None else max(0.02, random.gauss(0.09, 0.025))
    log("bot", "down", name)
    kb.press(key)
    time.sleep(hold)
    kb.release(key)
    log("bot", "up", name)


def type_char(ch):
    if ch == "\n":
        press(keyboard.Key.enter)
    else:
        press(ch)


def run(text, persona):
    global abort
    text = text.replace("\t", "    ")
    i, prev = 0, ""
    start = time.perf_counter()
    while i < len(text) and not abort:
        ch = text[i]
        elapsed = time.perf_counter() - start

        # Typo: wrong neighbor key, maybe keep typing, notice, backspace, fix.
        if ch.lower() in NEIGHBORS and random.random() < persona.typo_rate:
            wrong = random.choice(NEIGHBORS[ch.lower()])
            wrong = wrong.upper() if ch.isupper() else wrong
            extra = random.choice([0, 0, 1, 2])
            span = text[i + 1:i + 1 + extra]
            time.sleep(char_delay(persona, prev, wrong, elapsed))
            type_char(wrong)
            for c in span:
                time.sleep(char_delay(persona, wrong, c, elapsed))
                type_char(c)
            time.sleep(random.uniform(0.25, 0.6))          # notice the mistake
            for _ in range(1 + len(span)):
                press(keyboard.Key.backspace)
                time.sleep(random.uniform(0.05, 0.12))
            log("bot", "correction", ch, typo=wrong)
            continue  # retry the same index without a typo this time

        time.sleep(char_delay(persona, prev, ch, elapsed))
        type_char(ch)
        prev = ch
        i += 1

        if ch == "\n":
            if random.random() < 0.5:
                time.sleep(random.uniform(*persona.line_pause))
        elif ch == " " and random.random() < persona.word_pause_prob:
            time.sleep(random.uniform(0.15, 0.7))


def on_observed(key):
    global abort
    if key == keyboard.Key.f12:
        abort = True
        return False
    log("observed", "down", str(key))


def main():
    global log_file
    template = sys.argv[1] if len(sys.argv) > 1 else "template.py"
    persona = PERSONAS.get(sys.argv[2] if len(sys.argv) > 2 else "average")
    text = open(template, encoding="utf-8").read()

    log_file = open("session_log.jsonl", "w", encoding="utf-8")
    log("bot", "session_start", None, persona=persona.name, template=template)

    listener = keyboard.Listener(on_press=on_observed)
    listener.start()

    for n in range(5, 0, -1):
        print(f"Click into VS Code... starting in {n}")
        time.sleep(1)

    run(text, persona)
    log("bot", "session_end", None, aborted=abort)
    log_file.close()
    listener.stop()
    print("Done." if not abort else "Aborted.")


if __name__ == "__main__":
    main()