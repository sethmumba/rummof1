import keyword
import math
import random
import re
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from pynput.keyboard import Controller as KeyboardController, Key
from hm import human_move, human_click, human_reading_wander, human_scroll

keyboard = KeyboardController()

# --- FILE PATH SETUP ---
SCRIPT_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = SCRIPT_DIR / "template.py"

# --- WORKDAY SCHEDULE CONFIGURATION ---
STARTUP_DELAY_SECONDS = 5             # Buffer to switch to VS Code
WORK_SPRINT_MINUTES = (75, 105)       # Continuous work duration
LONG_BREAK_MINUTES = (12, 18)         # Full rest break away from desk
SHORT_BREAK_MINUTES = (1, 3)          # Quick water/stretch break
SHORT_BREAK_INTERVAL_LINES = (35, 60) # Trigger a short break every N lines

# --- COORDINATES (Calibrated for your VS Code layout) ---
TAB_TARGET_PY = (1245, 48)            # Position of target.py top tab
TAB_TEMPLATE_PY = (1334, 45)          # Position of template.py top tab

# --- KEYBOARD TOPOLOGY & BIOMETRICS ---
ADJACENT_KEYS = {
    'a': ['q', 'w', 's', 'z'], 'b': ['v', 'g', 'h', 'n'], 'c': ['x', 'd', 'f', 'v'],
    'd': ['s', 'e', 'r', 'f', 'x', 'c'], 'e': ['w', 's', 'd', 'r', '3', '4'],
    'f': ['d', 'r', 't', 'g', 'c', 'v'], 'g': ['f', 't', 'y', 'h', 'v', 'b'],
    'h': ['g', 'y', 'u', 'j', 'b', 'n'], 'i': ['u', 'j', 'k', 'o', '8', '9'],
    'j': ['h', 'u', 'i', 'k', 'n', 'm'], 'k': ['j', 'i', 'o', 'l', 'm'],
    'l': ['k', 'o', 'p'], 'm': ['n', 'j', 'k'], 'n': ['b', 'h', 'j', 'm'],
    'o': ['i', 'k', 'l', 'p', '9', '0'], 'p': ['o', 'l', '0'],
    'r': ['e', 'd', 'f', 't', '4', '5'], 's': ['a', 'w', 'e', 'd', 'x', 'z'],
    't': ['r', 'f', 'g', 'y', '5', '6'], 'u': ['y', 'h', 'j', 'i', '7', '8'],
    'v': ['c', 'f', 'g', 'b'], 'w': ['q', 'a', 's', 'e', '2', '3'],
    'x': ['z', 's', 'd', 'c'], 'y': ['t', 'g', 'h', 'u', '6', '7'],
    'z': ['a', 's', 'x']
}

COMMON_BUILTINS = {
    "print", "len", "range", "str", "int", "float", "list", "dict",
    "set", "tuple", "sum", "min", "max", "open", "isinstance", "enumerate",
    "zip", "super", "self", "cls", "True", "False", "None"
}

TOKEN_REGEX = re.compile(
    r"[a-zA-Z_][a-zA-Z0-9_]*"    # Identifiers, keywords
    r"|\d+(?:\.\d+)?"             # Numbers
    r"|[ \t]+"                    # Whitespace
    r"|\n"                        # Newlines
    r"|[^\w\s]"                   # Operators & punctuation
)

def tap_key(key):
    """Presses and releases a key with realistic finger contact hold-time."""
    hold = max(0.038, random.lognormvariate(math.log(0.075), 0.20))
    keyboard.press(key)
    time.sleep(hold)
    keyboard.release(key)

def restore_editor_caret():
    """Resets the text insertion caret cleanly to the end of the file."""
    keyboard.press(Key.ctrl)
    time.sleep(random.uniform(0.05, 0.09))
    tap_key(Key.end)
    time.sleep(random.uniform(0.04, 0.08))
    keyboard.release(Key.ctrl)
    time.sleep(random.uniform(0.08, 0.15))

def human_countdown_break(minutes, label="Break"):
    """Simulates realistic human away-time with progress logging."""
    total_seconds = int(minutes * 60)
    end_time = datetime.now() + timedelta(seconds=total_seconds)
    print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Entering {label} for {minutes:.1f} minutes...")
    print(f"Resuming at approximately: {end_time.strftime('%H:%M:%S')}")

    elapsed = 0
    while elapsed < total_seconds:
        sleep_slice = min(15, total_seconds - elapsed)
        time.sleep(sleep_slice)
        elapsed += sleep_slice

    print(f"[{datetime.now().strftime('%H:%M:%S')}] {label} ended. Returning to desk...")
    time.sleep(random.uniform(2.5, 4.0))

def get_chunk_profile(token_text):
    pre_pause = 0.0
    speed_factor = 1.0

    if keyword.iskeyword(token_text):
        speed_factor = random.uniform(1.30, 1.55)
        if token_text in {"def", "class", "if", "elif", "while", "try", "except"}:
            pre_pause = random.uniform(0.60, 1.80)
        else:
            pre_pause = random.uniform(0.08, 0.20)
    elif token_text in COMMON_BUILTINS:
        speed_factor = random.uniform(1.20, 1.40)
        pre_pause = random.uniform(0.05, 0.15)
    elif token_text.isidentifier():
        parts = token_text.split("_")
        if len(parts) > 1:
            speed_factor = random.uniform(0.75, 0.95)
            pre_pause = random.uniform(0.25, 0.65)
        else:
            speed_factor = random.uniform(0.85, 1.05)
            pre_pause = random.uniform(0.12, 0.35)
    elif token_text in {":", "->", "==", "!=", "<=", ">="}:
        pre_pause = random.uniform(0.10, 0.25)
        speed_factor = 0.90

    return speed_factor, pre_pause

def run_workday_simulation():
    if not TEMPLATE_PATH.exists():
        print(f"Error: Could not find template file at: {TEMPLATE_PATH}")
        sys.exit(1)

    with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
        code_to_type = f.read()

    if not code_to_type.strip():
        print(f"Warning: {TEMPLATE_PATH} is empty. Write code and save it (Ctrl+S).")
        sys.exit(1)

    print("==================================================")
    print(f"Loaded {len(code_to_type)} characters from {TEMPLATE_PATH.name}.")
    print("==================================================")

    for remaining in range(STARTUP_DELAY_SECONDS, 0, -1):
        print(f"Starting in {remaining}s... Focus target.py in VS Code now!", end="\r")
        time.sleep(1)

    print("\n[Active] Simulation started! Ensuring cursor is at line 1...")

    restore_editor_caret()
    keyboard.press(Key.ctrl)
    tap_key(Key.home)
    keyboard.release(Key.ctrl)
    time.sleep(0.3)

    sprint_duration = random.uniform(*WORK_SPRINT_MINUTES)
    sprint_end_time = datetime.now() + timedelta(minutes=sprint_duration)
    print(f"Current Sprint: {sprint_duration:.1f} minutes. Scheduled break at {sprint_end_time.strftime('%H:%M:%S')}")

    lines_since_micro_break = 0
    next_micro_break_target = random.randint(*SHORT_BREAK_INTERVAL_LINES)
    lines_typed_total = 0

    base_wpm = random.uniform(64, 72)
    base_error_rate = 0.020  # Slightly reduced to ensure clean replication

    tokens = TOKEN_REGEX.findall(code_to_type)
    t_idx = 0
    total_tokens = len(tokens)

    while t_idx < total_tokens:
        now = datetime.now()

        if now >= sprint_end_time:
            break_duration = random.uniform(*LONG_BREAK_MINUTES)
            human_countdown_break(break_duration, label="Long Rest Break")
            human_click(TAB_TARGET_PY[0], TAB_TARGET_PY[1])
            time.sleep(random.uniform(0.4, 0.8))
            restore_editor_caret()
            keyboard.press(Key.ctrl)
            tap_key(Key.end)
            keyboard.release(Key.ctrl)

            sprint_duration = random.uniform(*WORK_SPRINT_MINUTES)
            sprint_end_time = datetime.now() + timedelta(minutes=sprint_duration)
            print(f"New Sprint started ({sprint_duration:.1f} min). Next break at {sprint_end_time.strftime('%H:%M:%S')}")

        if lines_since_micro_break >= next_micro_break_target:
            short_break = random.uniform(*SHORT_BREAK_MINUTES)
            human_countdown_break(short_break, label="Short Stretch Break")
            human_click(TAB_TARGET_PY[0], TAB_TARGET_PY[1])
            restore_editor_caret()
            keyboard.press(Key.ctrl)
            tap_key(Key.end)
            keyboard.release(Key.ctrl)
            lines_since_micro_break = 0
            next_micro_break_target = random.randint(*SHORT_BREAK_INTERVAL_LINES)

        base_wpm = max(48, min(82, base_wpm + random.uniform(-0.35, 0.35)))
        base_char_delay = 60.0 / (base_wpm * 5)

        token = tokens[t_idx]

        if token == "\n":
            tap_key(Key.enter)
            lines_since_micro_break += 1
            lines_typed_total += 1
            time.sleep(random.uniform(0.35, 0.90))

            if lines_typed_total % random.randint(15, 25) == 0:
                action = random.choice(['wander', 'scroll', 'check_reference'])
                if action == 'wander':
                    human_reading_wander(TAB_TARGET_PY[0], TAB_TARGET_PY[1] + 150, passes=1)
                    restore_editor_caret()
                    keyboard.press(Key.ctrl)
                    tap_key(Key.end)
                    keyboard.release(Key.ctrl)
                elif action == 'scroll':
                    human_scroll(clicks=random.randint(1, 3), direction='up')
                    time.sleep(random.uniform(0.4, 0.8))
                    human_scroll(clicks=random.randint(1, 3), direction='down')
                    restore_editor_caret()
                    keyboard.press(Key.ctrl)
                    tap_key(Key.end)
                    keyboard.release(Key.ctrl)
                elif action == 'check_reference':
                    human_click(TAB_TEMPLATE_PY[0], TAB_TEMPLATE_PY[1])
                    time.sleep(random.uniform(1.0, 2.0))
                    human_reading_wander(TAB_TEMPLATE_PY[0], TAB_TEMPLATE_PY[1] + 100, passes=1)
                    human_click(TAB_TARGET_PY[0], TAB_TARGET_PY[1])
                    time.sleep(random.uniform(0.3, 0.6))
                    restore_editor_caret()
                    keyboard.press(Key.ctrl)
                    tap_key(Key.end)
                    keyboard.release(Key.ctrl)

            t_idx += 1
            continue

        elif "\t" in token:
            for _ in token:
                tap_key(Key.tab)
                time.sleep(random.uniform(0.08, 0.15))
            t_idx += 1
            continue

        elif token.isspace():
            for _ in token:
                tap_key(" ")
                time.sleep(random.uniform(0.03, 0.08))
            t_idx += 1
            continue

        speed_factor, pre_pause = get_chunk_profile(token)
        if pre_pause > 0:
            time.sleep(pre_pause)

        token_char_delay = base_char_delay / speed_factor

        # --- ROBUST CHARACTER TYPING WITH PROPER TYPO CORRECTION ---
        c_idx = 0
        while c_idx < len(token):
            char = token[c_idx]

            # Intentional adjacent typo simulation
            if (
                char.lower() in ADJACENT_KEYS
                and random.random() < base_error_rate
                and char not in ('\n', '\t', ' ')
            ):
                wrong_char = random.choice(ADJACENT_KEYS[char.lower()])
                tap_key(wrong_char)

                # Overshoot 0-1 characters
                overshoot_count = random.choices([0, 1], weights=[0.7, 0.3])[0]
                overshot_chars = []
                for o in range(1, overshoot_count + 1):
                    if c_idx + o < len(token):
                        nxt_c = token[c_idx + o]
                        tap_key(nxt_c)
                        overshot_chars.append(nxt_c)
                        time.sleep(random.uniform(0.06, 0.12))

                # Realization delay
                time.sleep(random.uniform(0.18, 0.38))

                # Erase mistakes (wrong char + overshot chars)
                total_to_delete = 1 + len(overshot_chars)
                for _ in range(total_to_delete):
                    tap_key(Key.backspace)
                    time.sleep(random.uniform(0.07, 0.13))

                time.sleep(random.uniform(0.10, 0.22))

                # FIXED: Now correctly re-types from the current character onward so nothing is skipped!
                correct_slice = token[c_idx:]
                for corr_char in correct_slice:
                    tap_key(corr_char)
                    time.sleep(max(0.022, random.lognormvariate(math.log(token_char_delay), 0.28)))
                break  # Token is now fully typed out correctly via slice

            if char == "_" and c_idx > 0:
                time.sleep(random.uniform(0.08, 0.20))

            tap_key(char)
            flight_time = random.lognormvariate(math.log(token_char_delay), 0.28)
            time.sleep(max(0.022, flight_time))
            c_idx += 1

        t_idx += 1

    print("\n[Completed] Entire template codebase typed cleanly without mangling.")

if __name__ == "__main__":
    run_workday_simulation()