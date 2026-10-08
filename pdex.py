"""
AI Training Data Entry Simulation Engine: CSV to Excel
------------------------------------------------------
A modular human typing simulator that:
- Reads and parses tabular data from a CSV file.
- Types cell-by-cell into Microsoft Excel using realistic human keystrokes.
- Formats headers automatically (Bolds them).
- ALWAYS Autofits column width (Alt+H, O, I) for EVERY cell so data never overlaps.
- Applies smart, context-aware formatting dynamically based on cell contents.
- Simulates realistic "chunk reading" (memorizes a row, types it fast, pauses at the end).
- Models a 15% chance to "glance back" at the source document mid-row.
- Models biomechanical finger movement, fatigue, drift, and thinking pauses.
- Simulates realistic typos, adjacent-key slips, transpositions, and backspace repairs.
- Maintains crash-resilient resumable state and ALWAYS resumes from Column A of the current row.
- Simulates Application Switching (Alt+Tab), Random Mouse Movements, and Scrolling.
- Injects Macro-Interruptions (Organic Chaos) to break statistical timing clusters.

Install:
    pip install pynput pywin32
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import math
import random
import statistics
import sys
import time
import uuid
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Optional

# Optional Windows/execution dependencies
try:
    from pynput.keyboard import Controller as KeyboardController, Key
    from pynput.mouse import Controller as MouseController, Button
except Exception:
    KeyboardController = None
    MouseController = None
    Key = None
    Button = None

try:
    from hm import human_click
except Exception:
    human_click = None


# ============================================================
# LOGGING SETUP
# ============================================================

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("ExcelTypingSim")


# ============================================================
# CONFIGURATION
# ============================================================

if getattr(sys, 'frozen', False):
    # Running as a compiled PyInstaller executable
    SCRIPT_DIR = Path(sys.executable).parent
else:
    # Running as a normal Python script
    SCRIPT_DIR = Path(__file__).resolve().parent

CSV_PATH = SCRIPT_DIR / "input.csv"

OUTPUT_DIR = Path(r"C:\typing_sim\simulation_output_excel")
STATE_FILE = OUTPUT_DIR / "typing_state.json"
EVENT_LOG_FILE = OUTPUT_DIR / "events.jsonl"
SESSION_SUMMARY_FILE = OUTPUT_DIR / "session_summary.json"
CALIBRATION_REPORT_FILE = OUTPUT_DIR / "calibration_report.json"

REFERENCE_DATA_PATH: Optional[Path] = None
ENABLE_EMPIRICAL_CALIBRATION = True
CALIBRATION_MAX_ROWS = 250_000
CALIBRATION_TRIM_FRACTION = 0.01

EXECUTE_IN_APPLICATION = True
RANDOM_SEED: Optional[int] = 20261004
PROFILE_NAME = "average"

STARTUP_DELAY_SECONDS = 60
WORK_SPRINT_MINUTES = (75, 105)
LONG_BREAK_MINUTES = (12, 18)
SHORT_BREAK_MINUTES = (1, 3)
SHORT_BREAK_INTERVAL_ROWS = (25, 45)

# X and Y coordinates to click to ensure Excel is focused before typing
SAFE_EXCEL_CELL_CLICK = (599, 16)  
MAX_SESSION_HOURS = 12


# ============================================================
# DATA MODELS
# ============================================================

@dataclass
class HumanProfile:
    name: str
    base_wpm: float
    wpm_stddev: float
    min_wpm: float
    max_wpm: float

    base_error_rate: float
    correction_probability: float
    delayed_correction_probability: float
    omission_probability: float
    duplication_probability: float
    transposition_probability: float
    punctuation_error_probability: float
    capitalization_error_probability: float
    space_error_probability: float

    pause_probability: float
    word_pause_ms: tuple[float, float]
    cell_pause_ms: tuple[float, float]
    row_pause_ms: tuple[float, float]
    thinking_pause_ms: tuple[float, float]

    key_interval_sigma: float
    correction_delay_ms: tuple[float, float]

    fatigue_rate: float
    fatigue_error_multiplier: float
    recovery_rate: float

    uncorrected_probability: float


# Profiles adjusted for "Chunk Reading"
PROFILES = {
    "average": HumanProfile(
        name="average",
        base_wpm=48, wpm_stddev=7, min_wpm=32, max_wpm=72,
        base_error_rate=0.018,
        correction_probability=0.82,
        delayed_correction_probability=0.13,
        omission_probability=0.22,
        duplication_probability=0.16,
        transposition_probability=0.22,
        punctuation_error_probability=0.15,
        capitalization_error_probability=0.16,
        space_error_probability=0.13,
        pause_probability=0.055,
        word_pause_ms=(60, 260),
        cell_pause_ms=(200, 500),   
        row_pause_ms=(2500, 5000),  
        thinking_pause_ms=(1500, 4000),
        key_interval_sigma=0.28,
        correction_delay_ms=(120, 900),
        fatigue_rate=0.005,
        fatigue_error_multiplier=0.75,
        recovery_rate=0.15,
        uncorrected_probability=0.035,
    ),
}


@dataclass
class SimulationConfig:
    seed: Optional[int]
    profile_name: str
    execute: bool
    output_dir: Path
    event_log: Path
    state_file: Path
    summary_file: Path

    def validate(self) -> None:
        if self.profile_name not in PROFILES:
            raise ValueError(
                f"Unknown profile {self.profile_name!r}. "
                f"Choose from: {', '.join(PROFILES)}"
            )


@dataclass
class CellItem:
    row_idx: int
    col_idx: int
    text: str
    is_last_in_row: bool


@dataclass
class Event:
    event_id: int
    session_id: str
    timestamp: str
    elapsed_ms: float
    event_type: str

    intended: Optional[str] = None
    actual: Optional[str] = None
    latency_ms: Optional[float] = None

    error_type: Optional[str] = None
    corrected: Optional[bool] = None
    correction_delay_ms: Optional[float] = None

    row_idx: Optional[int] = None
    col_idx: Optional[int] = None
    wpm: Optional[float] = None
    fatigue: Optional[float] = None
    metadata: dict[str, Any] = field(default_factory=dict)


# ============================================================
# KEYBOARD GEOMETRY & BIOMECHANICAL MODEL
# ============================================================

KEY_POSITIONS = {
    "q": (0.0, 0.0), "w": (1.0, 0.0), "e": (2.0, 0.0),
    "r": (3.0, 0.0), "t": (4.0, 0.0), "y": (5.0, 0.0),
    "u": (6.0, 0.0), "i": (7.0, 0.0), "o": (8.0, 0.0),
    "p": (9.0, 0.0),
    "a": (0.25, 1.0), "s": (1.25, 1.0), "d": (2.25, 1.0),
    "f": (3.25, 1.0), "g": (4.25, 1.0), "h": (5.25, 1.0),
    "j": (6.25, 1.0), "k": (7.25, 1.0), "l": (8.25, 1.0),
    "z": (0.75, 2.0), "x": (1.75, 2.0), "c": (2.75, 2.0),
    "v": (3.75, 2.0), "b": (4.75, 2.0), "n": (5.75, 2.0),
    "m": (6.75, 2.0),
    "1": (0.0, -1.0), "2": (1.0, -1.0), "3": (2.0, -1.0),
    "4": (3.0, -1.0), "5": (4.0, -1.0), "6": (5.0, -1.0),
    "7": (6.0, -1.0), "8": (7.0, -1.0), "9": (8.0, -1.0),
    "0": (9.0, -1.0),
}

FINGER = {
    **{k: "L_pinky" for k in "qaz1"},
    **{k: "L_ring" for k in "wsx2"},
    **{k: "L_middle" for k in "edc3"},
    **{k: "L_index" for k in "rfvtgb4"},
    **{k: "R_index" for k in "yhnujm67"},
    **{k: "R_middle" for k in "ik8"},
    **{k: "R_ring" for k in "ol9"},
    **{k: "R_pinky" for k in "p0"},
}

def key_distance(a: str, b: str) -> float:
    a = a.lower()
    b = b.lower()
    if a not in KEY_POSITIONS or b not in KEY_POSITIONS:
        return 4.0
    ax, ay = KEY_POSITIONS[a]
    bx, by = KEY_POSITIONS[b]
    return math.hypot(ax - bx, ay - by)

def nearby_keys(char: str, radius: float = 1.65) -> list[str]:
    c = char.lower()
    if c not in KEY_POSITIONS:
        return []
    return [k for k in KEY_POSITIONS if k != c and key_distance(c, k) <= radius]


# ============================================================
# CSV PARSER
# ============================================================

class CSVDocumentParser:
    def __init__(self, csv_path: Path):
        self.csv_path = csv_path

    def parse(self) -> list[CellItem]:
        if not self.csv_path.exists():
            raise FileNotFoundError(f"CSV file not found: {self.csv_path}")

        cells: list[CellItem] = []
        with self.csv_path.open("r", encoding="utf-8-sig", newline="") as f:
            sample = f.read(8192)
            f.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample)
            except csv.Error:
                dialect = csv.excel

            reader = csv.reader(f, dialect=dialect)
            for row_idx, row in enumerate(reader):
                if not row:
                    continue
                num_cols = len(row)
                for col_idx, value in enumerate(row):
                    cells.append(
                        CellItem(
                            row_idx=row_idx,
                            col_idx=col_idx,
                            text=str(value),
                            is_last_in_row=(col_idx == num_cols - 1),
                        )
                    )
        return cells


# ============================================================
# HUMAN BEHAVIOR MODEL
# ============================================================

class HumanBehaviorModel:
    def __init__(self, profile: HumanProfile, rng: random.Random):
        self.profile = profile
        self.rng = rng
        self.fatigue = 0.0
        self.current_wpm = self._initial_wpm()

    def _initial_wpm(self) -> float:
        value = self.rng.gauss(self.profile.base_wpm, self.profile.wpm_stddev)
        return max(self.profile.min_wpm, min(self.profile.max_wpm, value))

    def recover(self, minutes: float) -> None:
        self.fatigue *= max(0.0, math.exp(-self.profile.recovery_rate * minutes))
        self.current_wpm = min(
            self.profile.max_wpm,
            self.current_wpm + self.profile.wpm_stddev * 0.25,
        )

    def advance(self, elapsed_minutes: float) -> None:
        self.fatigue = min(1.0, self.fatigue + elapsed_minutes * self.profile.fatigue_rate)
        drift = self.rng.gauss(0, self.profile.wpm_stddev * 0.025)
        target = self.profile.base_wpm * max(0.60, 1.0 - self.fatigue * 0.25)
        self.current_wpm = max(
            self.profile.min_wpm,
            min(self.profile.max_wpm, self.current_wpm + (target - self.current_wpm) * 0.03 + drift),
        )

    def error_probability(self, char: str) -> float:
        punc_mult = 1.5 if char in ".,;:!?()[]{}'\"-" else 1.0
        fatigue_mult = 1.0 + (self.fatigue * self.profile.fatigue_error_multiplier)
        return min(0.25, self.profile.base_error_rate * punc_mult * fatigue_mult)

    def base_char_delay(self) -> float:
        cps = max(1.0, self.current_wpm * 5.0 / 60.0)
        return 1.0 / cps

    def character_delay(self, previous: Optional[str], current: str) -> float:
        base = self.base_char_delay()
        multiplier = 1.0

        if previous and previous.lower() in FINGER and current.lower() in FINGER:
            if FINGER[previous.lower()] == FINGER[current.lower()]:
                multiplier += 0.10
            elif FINGER[previous.lower()].startswith(FINGER[current.lower()][0]):
                multiplier += 0.025
            else:
                multiplier -= 0.025

        if current in ".,;:!?":
            multiplier += 0.10

        delay = self.rng.lognormvariate(
            math.log(max(0.025, base * multiplier)),
            self.profile.key_interval_sigma,
        )
        return max(0.025, min(1.8, delay))

    def should_pause(self, current: str) -> bool:
        prob = self.profile.pause_probability
        if current in ".!?":
            prob *= 2.0
        return self.rng.random() < prob

    def pause_duration(self, reason: str) -> float:
        if reason == "word_boundary":
            lo, hi = self.profile.word_pause_ms
        elif reason == "cell_transition":
            lo, hi = self.profile.cell_pause_ms
        elif reason == "row_transition":
            lo, hi = self.profile.row_pause_ms
        else:
            lo, hi = self.profile.thinking_pause_ms

        midpoint = max(1.0, (lo + hi) / 2)
        val = self.rng.lognormvariate(math.log(midpoint), 0.45)
        return max(lo, min(hi * 2.5, val))

    def choose_error(self, char: str) -> Optional[str]:
        if self.rng.random() >= self.error_probability(char):
            return None

        choices: list[tuple[str, float]] = []
        if char.isalpha():
            choices.extend([
                ("substitution", 0.45),
                ("omission", self.profile.omission_probability),
                ("duplication", self.profile.duplication_probability),
                ("transposition", self.profile.transposition_probability),
                ("capitalization", self.profile.capitalization_error_probability),
            ])
        if char in ".,;:!?":
            choices.append(("punctuation", self.profile.punctuation_error_probability))
        if char == " ":
            choices.append(("space", self.profile.space_error_probability))

        names = [n for n, w in choices if w > 0]
        weights = [w for _, w in choices if w > 0]
        return self.rng.choices(names, weights=weights, k=1)[0] if names else "substitution"

    def correction_mode(self) -> str:
        r = self.rng.random()
        if r < self.profile.uncorrected_probability:
            return "none"
        if r < self.profile.uncorrected_probability + self.profile.delayed_correction_probability:
            return "delayed"
        if r < (self.profile.uncorrected_probability +
                self.profile.delayed_correction_probability +
                self.profile.correction_probability):
            return "immediate"
        return "none"


# ============================================================
# EVENT LOGGER & METRICS
# ============================================================

class EventLogger:
    def __init__(self, path: Path):
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if not self.path.exists():
            self.path.write_text("", encoding="utf-8")

    def write(self, event: Event) -> None:
        with self.path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(asdict(event), ensure_ascii=False) + "\n")


class Metrics:
    def __init__(self):
        self.events: list[Event] = []

    def add(self, event: Event) -> None:
        self.events.append(event)

    def summary(self) -> dict[str, Any]:
        if not self.events:
            return {"event_count": 0}

        key_events = [e for e in self.events if e.event_type == "key"]
        errors = [e for e in self.events if e.error_type]
        corrections = [e for e in self.events if e.event_type == "correction"]
        pauses = [e for e in self.events if e.event_type == "pause"]

        elapsed_ms = max(e.elapsed_ms for e in self.events)
        actual_chars = sum(1 for e in key_events if e.actual is not None)
        words = sum(1 for e in key_events if e.actual == " ")
        minutes = max(0.001, elapsed_ms / 60000)

        latencies = [e.latency_ms for e in key_events if e.latency_ms is not None]

        return {
            "event_count": len(self.events),
            "key_events": len(key_events),
            "characters": actual_chars,
            "duration_minutes": round(minutes, 3),
            "effective_wpm": round(words / minutes, 3),
            "error_count": len(errors),
            "error_rate": round(len(errors) / max(1, len(key_events)), 6),
            "correction_count": len(corrections),
            "pause_count": len(pauses),
        }


# ============================================================
# STATE MANAGEMENT
# ============================================================

class StateManager:
    def __init__(self, path: Path):
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def load(self) -> dict[str, Any]:
        if not self.path.exists():
            return {"cell_index": 0, "session_id": None, "seed": None}
        try:
            return json.loads(self.path.read_text(encoding="utf-8"))
        except Exception as exc:
            logger.warning(f"State file error: {exc}")
            return {"cell_index": 0, "session_id": None, "seed": None}

    def save(self, cell_index: int, session_id: str, seed: int) -> None:
        payload = {
            "cell_index": cell_index,
            "session_id": session_id,
            "seed": seed,
            "updated_at": datetime.now(timezone.utc).isoformat(),
        }
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        tmp.replace(self.path)

    def clear(self) -> None:
        if self.path.exists():
            self.path.unlink()


# ============================================================
# EXCEL / WINDOWS EXECUTOR
# ============================================================

class WindowsExcelExecutor:
    def __init__(self, enabled: bool):
        self.enabled = enabled
        self.keyboard = (
            KeyboardController() if enabled and KeyboardController is not None else None
        )
        self.mouse = (
            MouseController() if enabled and MouseController is not None else None
        )
        if enabled and self.keyboard is None:
            raise RuntimeError("Keyboard execution requested, but pynput is unavailable.")

    # --- Mouse & Navigation Integrations ---
    def random_mouse_nudge(self, rng: random.Random) -> None:
        if not self.enabled or not self.mouse:
            return
        current_x, current_y = self.mouse.position
        dx = rng.randint(-15, 15)
        dy = rng.randint(-15, 15)
        self.mouse.position = (current_x + dx, current_y + dy)

    def scroll_randomly(self, rng: random.Random) -> None:
        if not self.enabled or not self.mouse:
            return
        scroll_amount = rng.randint(2, 6)
        self.mouse.scroll(0, scroll_amount)
        time.sleep(rng.uniform(0.3, 0.8))
        self.mouse.scroll(0, -scroll_amount)
        time.sleep(rng.uniform(0.2, 0.5))
        self.random_mouse_nudge(rng)

    def switch_application(self, rng: random.Random, delay: float, scroll_source: bool = False) -> None:
        if not self.enabled:
            return
        
        self.safely_exit_ribbon()
        
        # Switch to the source document (via Z-order)
        self.keyboard.press(Key.alt)
        time.sleep(0.1)
        self.tap(Key.tab)
        time.sleep(0.1)
        self.keyboard.release(Key.alt)
        
        # Simulate reading the first half of the data
        time.sleep(delay / 2)
        
        # Optionally scroll down the source document to read more
        if scroll_source and self.mouse:
            scroll_amount = rng.randint(-6, -2)  # Negative values scroll down
            self.mouse.scroll(0, scroll_amount)
            # Pause again after scrolling to read the new text
            time.sleep(rng.uniform(0.5, 1.5))
            
        # Simulate reading the second half of the data
        time.sleep(delay / 2)
        
        # Switch back to Excel
        self.keyboard.press(Key.alt)
        time.sleep(0.1)
        self.tap(Key.tab)
        time.sleep(0.1)
        self.keyboard.release(Key.alt)
        time.sleep(0.5)

    def tap(self, key: Any, hold: float = 0.05) -> None:
        if not self.enabled:
            return
        self.keyboard.press(key)
        time.sleep(max(0.005, hold))
        self.keyboard.release(key)

    def _hold_and_tap(self, modifiers: list[Any], key: Any) -> None:
        if not self.enabled:
            return
        for mod in modifiers:
            self.keyboard.press(mod)
            time.sleep(0.05)
        self.tap(key, hold=0.05)
        for mod in reversed(modifiers):
            self.keyboard.release(mod)
            time.sleep(0.05)

    def safely_exit_ribbon(self) -> None:
        if not self.enabled:
            return
        self.tap(Key.esc)
        time.sleep(0.05)
        self.tap(Key.esc)
        time.sleep(0.05)

    def _alt_sequence(self, keys: list[str]) -> None:
        if not self.enabled:
            return
        self.safely_exit_ribbon()
        
        self.tap(Key.alt)
        time.sleep(0.35) 
        for k in keys:
            self.tap(k)
            time.sleep(0.2) 
            
        time.sleep(0.2)
        self.safely_exit_ribbon()

    def backspace(self) -> None:
        if self.enabled: self.tap(Key.backspace)

    def commit_in_place(self) -> None:
        self._hold_and_tap([Key.ctrl], Key.enter)
        time.sleep(0.1)

    def home(self) -> None:
        if self.enabled:
            self.tap(Key.home)
            time.sleep(0.1)

    def tab(self) -> None:
        if self.enabled: self.tap(Key.tab)

    def down(self) -> None:
        if self.enabled: self.tap(Key.down)

    def alt_enter(self) -> None:
        self._hold_and_tap([Key.alt], Key.enter)

    def toggle_bold(self) -> None:
        self._hold_and_tap([Key.ctrl], 'b')

    def align_center(self) -> None:
        self._alt_sequence(['h', 'a', 'c'])

    def align_right(self) -> None:
        self._alt_sequence(['h', 'a', 'r'])

    def wrap_text(self) -> None:
        self._alt_sequence(['h', 'w'])

    def all_borders(self) -> None:
        self._alt_sequence(['h', 'b', 'a'])

    def fill_color_default(self) -> None:
        self._alt_sequence(['h', 'h'])
        time.sleep(0.1)
        if self.enabled: self.tap(Key.enter)
        self.safely_exit_ribbon()

    def format_currency(self) -> None:
        self._hold_and_tap([Key.ctrl, Key.shift], '4')  

    def format_percentage(self) -> None:
        self._hold_and_tap([Key.ctrl, Key.shift], '5')  

    def format_date(self) -> None:
        self._hold_and_tap([Key.ctrl, Key.shift], '3')  

    def autofit_column(self) -> None:
        self._alt_sequence(['h', 'o', 'i'])

    def autofit_row(self) -> None:
        self._alt_sequence(['h', 'o', 'a'])

    def insert_filters(self) -> None:
        self._hold_and_tap([Key.ctrl, Key.shift], 'l')


# ============================================================
# SIMULATION ENGINE
# ============================================================

class ExcelTypingSimulation:
    def __init__(
        self,
        cells: list[CellItem],
        config: SimulationConfig,
        profile: HumanProfile,
        seed: int,
        session_id: str,
    ):
        self.cells = cells
        self.config = config
        self.profile = profile
        self.seed = seed
        self.session_id = session_id

        self.rng = random.Random(seed)
        self.behavior = HumanBehaviorModel(profile, self.rng)
        self.logger = EventLogger(config.event_log)
        self.metrics = Metrics()
        self.state = StateManager(config.state_file)
        self.executor = WindowsExcelExecutor(config.execute)

        self.event_id = 0
        self.elapsed_ms = 0.0
        self.rows_since_break = 0
        self.next_break_target = self.rng.randint(*SHORT_BREAK_INTERVAL_ROWS)

    def emit(self, event_type: str, **kwargs: Any) -> Event:
        self.event_id += 1
        event = Event(
            event_id=self.event_id,
            session_id=self.session_id,
            timestamp=datetime.now(timezone.utc).isoformat(),
            elapsed_ms=round(self.elapsed_ms, 3),
            event_type=event_type,
            wpm=round(self.behavior.current_wpm, 3),
            fatigue=round(self.behavior.fatigue, 6),
            **kwargs,
        )
        self.logger.write(event)
        self.metrics.add(event)
        return event

    def sleep_simulated(self, seconds: float) -> None:
        self.elapsed_ms += seconds * 1000
        if self.config.execute:
            time.sleep(seconds)
        self.behavior.advance(seconds / 60)

    def emit_pause(self, duration_ms: float, reason: str, **kwargs: Any) -> None:
        self.emit(
            "pause",
            metadata={"duration_ms": round(duration_ms, 3), "reason": reason},
            **kwargs,
        )
        self.sleep_simulated(duration_ms / 1000)

    def substitution(self, char: str) -> Optional[str]:
        candidates = nearby_keys(char)
        if not candidates:
            return None
        weights = [math.exp(-1.5 * key_distance(char, c)) for c in candidates]
        return self.rng.choices(candidates, weights=weights, k=1)[0]

    def process_error(self, char: str, error_type: str, cell: CellItem) -> tuple[str, bool]:
        mode = self.behavior.correction_mode()

        if error_type == "substitution":
            wrong = self.substitution(char)
            if wrong is None:
                return char, True

            self.emit(
                "key",
                intended=char,
                actual=wrong,
                error_type="substitution",
                row_idx=cell.row_idx,
                col_idx=cell.col_idx,
            )
            self.executor.tap(wrong)
            self.sleep_simulated(self.behavior.character_delay(None, wrong))

            if mode in ("immediate", "delayed"):
                delay = self.rng.uniform(*self.profile.correction_delay_ms)
                if mode == "delayed":
                    delay *= self.rng.uniform(1.5, 3.0)
                self.sleep_simulated(delay / 1000)
                self.executor.backspace()
                self.emit(
                    "correction",
                    intended=char,
                    actual=char,
                    corrected=True,
                    correction_delay_ms=delay,
                    row_idx=cell.row_idx,
                    col_idx=cell.col_idx,
                )
                self.executor.tap(char)
                self.sleep_simulated(self.behavior.character_delay(wrong, char))
                return char, True

            return wrong, True
        return char, True

    def type_cell(self, cell: CellItem) -> None:
        text = cell.text
        i = 0
        previous: Optional[str] = None

        while i < len(text):
            char = text[i]

            if char == "\n":
                self.executor.alt_enter()
                self.sleep_simulated(self.rng.uniform(0.2, 0.5))
                i += 1
                previous = "\n"
                continue

            error_type = self.behavior.choose_error(char)
            if error_type:
                actual, advance = self.process_error(char, error_type, cell)
                if actual:
                    previous = actual[-1]
                if advance:
                    i += 1
                continue

            delay = self.behavior.character_delay(previous, char)
            self.executor.tap(char)
            self.emit(
                "key",
                intended=char,
                actual=char,
                latency_ms=delay * 1000,
                row_idx=cell.row_idx,
                col_idx=cell.col_idx,
            )
            self.sleep_simulated(delay)

            # Randomly drift the mouse mid-word on long cells
            if self.rng.random() < 0.005:
                self.executor.random_mouse_nudge(self.rng)

            previous = char
            i += 1

    def apply_smart_formatting(self, cell: CellItem) -> None:
        text = cell.text.strip()
        if not text:
            return

        if cell.row_idx == 0:
            self.executor.align_center()
            time.sleep(0.2)
            self.executor.all_borders()
            time.sleep(0.2)
            self.executor.fill_color_default()
            
            if cell.is_last_in_row:
                self.executor.insert_filters()
            return

        if len(text) > 40 and " " in text:
            self.executor.wrap_text()
            self.executor.autofit_row()
        elif text.startswith("$") or text.startswith("€") or text.startswith("£"):
            self.executor.format_currency()
        elif text.endswith("%"):
            self.executor.format_percentage()
        elif "-" in text and len(text) >= 8 and sum(c.isdigit() for c in text) >= 6:
            self.executor.format_date()
        elif text.replace(".", "").replace(",", "").isdigit():
            self.executor.align_right()

    def run(self, start_index: int = 0) -> dict[str, Any]:
        session_started = datetime.now(timezone.utc)
        sprint_duration = self.rng.uniform(*WORK_SPRINT_MINUTES)
        sprint_start_ms = self.elapsed_ms

        total_cells = len(self.cells)
        logger.info(f"Session starting typing data loop...")

        for index in range(start_index, total_cells):
            if self.elapsed_ms / 3_600_000 >= MAX_SESSION_HOURS:
                logger.warning("Max session hours reached. Halting.")
                break

            cell = self.cells[index]

            sprint_elapsed = (self.elapsed_ms - sprint_start_ms) / 1000
            if sprint_elapsed >= sprint_duration * 60:
                duration = self.rng.uniform(*LONG_BREAK_MINUTES)
                logger.info(f"[LONG BREAK] Resting for {duration:.1f} minutes...")
                self.emit("break", metadata={"type": "long", "duration": duration})
                self.sleep_simulated(duration * 60)
                self.behavior.recover(duration)
                sprint_duration = self.rng.uniform(*WORK_SPRINT_MINUTES)
                sprint_start_ms = self.elapsed_ms

            # Injected Macro-Interruption (Organic Chaos)
            # 0.5% chance per cell to stop and walk away for 2 to 15 minutes
            if self.rng.random() < 0.005:
                interrupt_duration = self.rng.uniform(2 * 60, 15 * 60)
                logger.info(f"[ORGANIC CHAOS] Sudden interruption! Pausing for {interrupt_duration/60:.1f} min...")
                self.emit("break", metadata={"type": "organic_interruption", "duration": interrupt_duration})
                self.sleep_simulated(interrupt_duration)
                self.behavior.recover(interrupt_duration / 60)

            if cell.row_idx == 0:
                self.executor.toggle_bold()

            self.executor.safely_exit_ribbon()
            self.type_cell(cell)
            self.executor.commit_in_place()
            self.executor.autofit_column()
            self.apply_smart_formatting(cell)

            if cell.is_last_in_row:
                self.executor.safely_exit_ribbon()
                self.executor.down()
                self.executor.home()

                # Navigation Diversity: Occasional random scrolling between rows
                if self.rng.random() < 0.08:
                    logger.info(f"Row {cell.row_idx+1}: Simulating reading scroll/mouse refocus...")
                    self.executor.scroll_randomly(self.rng)

                self.rows_since_break += 1
                pause_dur = self.behavior.pause_duration("row_transition")
                self.emit_pause(pause_dur, "row_transition", row_idx=cell.row_idx, col_idx=cell.col_idx)

                if self.rows_since_break >= self.next_break_target:
                    brk_duration = self.rng.uniform(*SHORT_BREAK_MINUTES)
                    logger.info(f"[SHORT BREAK] Resting for {brk_duration:.1f} min...")
                    self.sleep_simulated(brk_duration * 60)
                    self.behavior.recover(brk_duration)
                    self.rows_since_break = 0
                    self.next_break_target = self.rng.randint(*SHORT_BREAK_INTERVAL_ROWS)
            else:
                self.executor.safely_exit_ribbon()
                self.executor.tab()
                
                # Accidental mouse bump
                if self.rng.random() < 0.05:
                    self.executor.random_mouse_nudge(self.rng)
                
                pause_dur = self.behavior.pause_duration("cell_transition")
                reason = "cell_transition"

                # Application Switching logic tied into the "glance back" probability
                if self.rng.random() < 0.15:
                    # INCREASED DELAY: 3.5 to 9 seconds to simulate actual reading
                    glance_pause = self.rng.uniform(3500, 9000)
                    logger.info(f"Row {cell.row_idx+1}: Checking source document...")
                    
                    if self.rng.random() < 0.40:
                        # 25% chance to scroll down in the source document while reading
                        needs_scroll = self.rng.random() < 0.25
                        self.executor.switch_application(
                            rng=self.rng, 
                            delay=glance_pause / 1000, 
                            scroll_source=needs_scroll
                        )
                        reason = "alt_tab_source_check"
                    else:
                        pause_dur += glance_pause
                        reason = "glance_back_at_source"

                self.emit_pause(pause_dur, reason, row_idx=cell.row_idx, col_idx=cell.col_idx)

            self.state.save(index + 1, self.session_id, self.seed)

        summary = self.metrics.summary()
        self.config.summary_file.write_text(json.dumps(summary, indent=2), encoding="utf-8")
        return summary


# ============================================================
# MAIN
# ============================================================

def build_cli() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Human CSV-to-Excel typing simulator")
    parser.add_argument("--csv", type=Path, default=CSV_PATH, help="Path to input CSV file.")
    parser.add_argument("--profile", choices=sorted(PROFILES), default=PROFILE_NAME)
    parser.add_argument("--seed", type=int, default=RANDOM_SEED)
    parser.add_argument("--offline", action="store_true", help="Generate events without typing.")
    return parser

def main() -> None:
    args = build_cli().parse_args()
    execute = EXECUTE_IN_APPLICATION and not args.offline

    config = SimulationConfig(
        seed=args.seed,
        profile_name=args.profile,
        execute=execute,
        output_dir=OUTPUT_DIR,
        event_log=EVENT_LOG_FILE,
        state_file=STATE_FILE,
        summary_file=SESSION_SUMMARY_FILE,
    )
    config.validate()
    config.output_dir.mkdir(parents=True, exist_ok=True)

    seed = config.seed or random.SystemRandom().randrange(0, 2**63 - 1)
    state_manager = StateManager(config.state_file)
    state = state_manager.load()

    session_id = state.get("session_id") or str(uuid.uuid4())
    start_index = int(state.get("cell_index", 0))

    profile = PROFILES[config.profile_name]
    parser = CSVDocumentParser(args.csv)
    cells = parser.parse()
    
    if 0 < start_index < len(cells):
        target_row = cells[start_index].row_idx
        while start_index > 0 and cells[start_index - 1].row_idx == target_row:
            start_index -= 1
            
    if config.execute:
        for remaining in range(STARTUP_DELAY_SECONDS, 0, -1):
            print(f"Starting in {remaining}s... Focus the correct cell in Excel!", end="\r", flush=True)
            time.sleep(1)
        print(" " * 65, end="\r")

    simulation = ExcelTypingSimulation(
        cells=cells, config=config, profile=profile, seed=seed, session_id=session_id
    )

    try:
        simulation.run(start_index=start_index)
        if state_manager.path.exists():
            state_manager.clear()
    except KeyboardInterrupt:
        logger.warning("\nInterrupted by user. State preserved.")
        raise

if __name__ == "__main__":
    main()