import json
import keyword
import math
import random
import re
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
import pymupdf  
from pynput.keyboard import Controller as KeyboardController, Key
from hm import human_move, human_click, human_reading_wander, human_scroll

keyboard = KeyboardController()

# --- FILE PATH SETUP ---
SCRIPT_DIR = Path(__file__).resolve().parent
PDF_PATH = SCRIPT_DIR / "input.pdf"
STATE_FILE = SCRIPT_DIR / "typing_state.json"  # File to track progress

# --- WORKDAY SCHEDULE CONFIGURATION ---
STARTUP_DELAY_SECONDS = 60            
WORK_SPRINT_MINUTES = (75, 105)       
LONG_BREAK_MINUTES = (12, 18)         
SHORT_BREAK_MINUTES = (1, 3)          
SHORT_BREAK_INTERVAL_LINES = (35, 60) 

# --- COORDINATES (Word/Docs Safe) ---
SAFE_TITLE_BAR_CLICK = (600, 15)  

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

def tap_key(key):
    hold = max(0.038, random.lognormvariate(math.log(0.075), 0.20))
    keyboard.press(key)
    time.sleep(hold)
    keyboard.release(key)

def toggle_bold():
    keyboard.press(Key.ctrl)
    time.sleep(random.uniform(0.05, 0.08))
    tap_key('b')
    time.sleep(random.uniform(0.04, 0.07))
    keyboard.release(Key.ctrl)
    time.sleep(random.uniform(0.15, 0.30))

def human_countdown_break(minutes, label="Break"):
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

def extract_pdf_spans(pdf_path):
    doc = pymupdf.open(pdf_path)
    content_stream = []

    for page_num, page in enumerate(doc):
        blocks = page.get_text("dict")["blocks"]
        for b in blocks:
            if "lines" not in b:
                continue
            for line in b["lines"]:
                for span in line["spans"]:
                    text = span["text"]
                    flags = span["flags"]
                    font_name = span["font"].lower()
                    
                    is_bold = bool(flags & 16) or "bold" in font_name or "heavy" in font_name or "black" in font_name
                    
                    if text:
                        content_stream.append({"text": text, "is_bold": is_bold})
                content_stream.append({"text": "\n", "is_bold": False})
        if page_num < len(doc) - 1:
            content_stream.append({"text": "\n\n", "is_bold": False})

    return content_stream

def load_state():
    """Loads the last processed span index if a state file exists."""
    if STATE_FILE.exists():
        try:
            with open(STATE_FILE, "r") as f:
                data = json.load(f)
                return data.get("span_idx", 0)
        except Exception as e:
            print(f"Warning: Could not read state file ({e}). Starting from beginning.")
    return 0

def save_state(span_idx):
    """Saves the current span index to a file to allow resuming."""
    with open(STATE_FILE, "w") as f:
        json.dump({"span_idx": span_idx}, f)

def run_pdf_simulation():
    if not PDF_PATH.exists():
        print(f"Error: Could not find PDF file at: {PDF_PATH}")
        sys.exit(1)

    print("Extracting text and formatting properties from PDF...")
    pdf_spans = extract_pdf_spans(PDF_PATH)
    
    if not pdf_spans:
        print(f"Warning: {PDF_PATH.name} yielded no readable text content.")
        sys.exit(1)

    total_spans = len(pdf_spans)
    span_idx = load_state()

    print("==================================================")
    print(f"Successfully loaded {total_spans} text segments from {PDF_PATH.name}.")
    if span_idx > 0:
        print(f"RESUMING FROM SAVED STATE: Segment {span_idx} of {total_spans}.")
        print("CRITICAL: Place your cursor exactly where the text left off, and ensure BOLD is toggled OFF in your editor before the timer ends.")
    else:
        print("Starting a fresh document simulation.")
    print("==================================================")

    for remaining in range(STARTUP_DELAY_SECONDS, 0, -1):
        print(f"Starting in {remaining}s... Focus your word processor window now!", end="\r")
        time.sleep(1)

    print("\n[Active] Simulation started! Clicking application title bar...")
    human_click(SAFE_TITLE_BAR_CLICK[0], SAFE_TITLE_BAR_CLICK[1])
    time.sleep(0.3)

    sprint_duration = random.uniform(*WORK_SPRINT_MINUTES)
    sprint_end_time = datetime.now() + timedelta(minutes=sprint_duration)
    print(f"Current Sprint: {sprint_duration:.1f} minutes. Scheduled break at {sprint_end_time.strftime('%H:%M:%S')}")

    lines_since_micro_break = 0
    next_micro_break_target = random.randint(*SHORT_BREAK_INTERVAL_LINES)
    lines_typed_total = 0

    base_wpm = random.uniform(64, 72)
    base_error_rate = 0.015
    currently_bold = False

    while span_idx < total_spans:
        now = datetime.now()

        if now >= sprint_end_time:
            if currently_bold:
                toggle_bold()
                currently_bold = False
            break_duration = random.uniform(*LONG_BREAK_MINUTES)
            human_countdown_break(break_duration, label="Long Rest Break")
            human_click(SAFE_TITLE_BAR_CLICK[0], SAFE_TITLE_BAR_CLICK[1])

            sprint_duration = random.uniform(*WORK_SPRINT_MINUTES)
            sprint_end_time = datetime.now() + timedelta(minutes=sprint_duration)
            print(f"New Sprint started ({sprint_duration:.1f} min). Next break at {sprint_end_time.strftime('%H:%M:%S')}")

        if lines_since_micro_break >= next_micro_break_target:
            if currently_bold:
                toggle_bold()
                currently_bold = False
            short_break = random.uniform(*SHORT_BREAK_MINUTES)
            human_countdown_break(short_break, label="Short Stretch Break")
            human_click(SAFE_TITLE_BAR_CLICK[0], SAFE_TITLE_BAR_CLICK[1])
            lines_since_micro_break = 0
            next_micro_break_target = random.randint(*SHORT_BREAK_INTERVAL_LINES)

        base_wpm = max(48, min(82, base_wpm + random.uniform(-0.35, 0.35)))
        base_char_delay = 60.0 / (base_wpm * 5)

        span_item = pdf_spans[span_idx]
        text_chunk = span_item["text"]
        target_bold = span_item["is_bold"]

        if target_bold != currently_bold:
            toggle_bold()
            currently_bold = target_bold

        if text_chunk == "\n" or text_chunk == "\n\n":
            for char in text_chunk:
                if char == "\n":
                    keyboard.press(Key.shift)
                    tap_key(Key.enter)
                    keyboard.release(Key.shift)
                    
                    lines_since_micro_break += 1
                    lines_typed_total += 1
                    time.sleep(random.uniform(0.35, 0.90))

                    if lines_typed_total % random.randint(15, 25) == 0:
                        action = random.choice(['wander', 'scroll'])
                        if action == 'wander':
                            human_reading_wander(SAFE_TITLE_BAR_CLICK[0], SAFE_TITLE_BAR_CLICK[1] + 150, passes=1)
                            human_click(SAFE_TITLE_BAR_CLICK[0], SAFE_TITLE_BAR_CLICK[1])
                        elif action == 'scroll':
                            human_scroll(clicks=random.randint(1, 3), direction='up')
                            time.sleep(random.uniform(0.4, 0.8))
                            human_scroll(clicks=random.randint(1, 3), direction='down')
                            human_click(SAFE_TITLE_BAR_CLICK[0], SAFE_TITLE_BAR_CLICK[1])

            span_idx += 1
            save_state(span_idx)
            continue

        c_idx = 0
        while c_idx < len(text_chunk):
            char = text_chunk[c_idx]

            if (
                char.lower() in ADJACENT_KEYS
                and random.random() < base_error_rate
                and char not in ('\n', '\t', ' ')
            ):
                wrong_char = random.choice(ADJACENT_KEYS[char.lower()])
                tap_key(wrong_char)
                time.sleep(random.uniform(0.18, 0.38))
                tap_key(Key.backspace)
                time.sleep(random.uniform(0.10, 0.22))

            if char == "_" and c_idx > 0:
                time.sleep(random.uniform(0.08, 0.20))

            tap_key(char)
            flight_time = random.lognormvariate(math.log(base_char_delay), 0.28)
            time.sleep(max(0.022, flight_time))
            c_idx += 1

        span_idx += 1
        save_state(span_idx)

    if currently_bold:
        toggle_bold()
        
    # Clean up the state file once the entire document is successfully completed
    if STATE_FILE.exists():
        STATE_FILE.unlink()

    print("\n[Completed] Entire PDF document typed cleanly with full routine and formatting retained.")

if __name__ == "__main__":
    run_pdf_simulation()