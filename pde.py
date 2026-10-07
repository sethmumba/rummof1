"""
AI Training Data Entry Simulation Engine: CSV to Excel
------------------------------------------------------
A modular human typing simulator that:
- Reads and parses tabular data from a CSV file.
- Types cell-by-cell into Microsoft Excel using realistic human keystrokes.
- Navigates via Tab (next column) and Enter (next row).
- Models biomechanical finger movement, fatigue, drift, and thinking pauses.
- Simulates realistic typos, adjacent-key slips, transpositions, and backspace repairs.
- Calibrates against empirical keystroke datasets (e.g., CMU dataset).
- Maintains crash-resilient resumable state and structured event logs.

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
except Exception:
    KeyboardController = None
    Key = None

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

STARTUP_DELAY_SECONDS = 25
WORK_SPRINT_MINUTES = (75, 105)
LONG_BREAK_MINUTES = (12, 18)
SHORT_BREAK_MINUTES = (1, 3)
SHORT_BREAK_INTERVAL_ROWS = (25, 45)

SAFE_EXCEL_CELL_CLICK = (559, 16)
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


PROFILES = {
    "fast_accurate": HumanProfile(
        name="fast_accurate",
        base_wpm=72, wpm_stddev=4, min_wpm=55, max_wpm=95,
        base_error_rate=0.008,
        correction_probability=0.90,
        delayed_correction_probability=0.07,
        omission_probability=0.18,
        duplication_probability=0.12,
        transposition_probability=0.18,
        punctuation_error_probability=0.10,
        capitalization_error_probability=0.10,
        space_error_probability=0.08,
        pause_probability=0.025,
        word_pause_ms=(45, 180),
        cell_pause_ms=(80, 220),
        row_pause_ms=(300, 700),
        thinking_pause_ms=(450, 1800),
        key_interval_sigma=0.20,
        correction_delay_ms=(100, 600),
        fatigue_rate=0.003,
        fatigue_error_multiplier=0.45,
        recovery_rate=0.18,
        uncorrected_probability=0.015,
    ),
    "average": HumanProfile(
        name="average",
        base_wpm=58, wpm_stddev=7, min_wpm=38, max_wpm=82,
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
        cell_pause_ms=(120, 400),
        row_pause_ms=(500, 1200),
        thinking_pause_ms=(600, 2400),
        key_interval_sigma=0.28,
        correction_delay_ms=(120, 900),
        fatigue_rate=0.005,
        fatigue_error_multiplier=0.75,
        recovery_rate=0.15,
        uncorrected_probability=0.035,
    ),
    "slow_careful": HumanProfile(
        name="slow_careful",
        base_wpm=43, wpm_stddev=5, min_wpm=28, max_wpm=62,
        base_error_rate=0.009,
        correction_probability=0.90,
        delayed_correction_probability=0.10,
        omission_probability=0.14,
        duplication_probability=0.10,
        transposition_probability=0.16,
        punctuation_error_probability=0.09,
        capitalization_error_probability=0.10,
        space_error_probability=0.08,
        pause_probability=0.095,
        word_pause_ms=(100, 380),
        cell_pause_ms=(180, 600),
        row_pause_ms=(800, 2000),
        thinking_pause_ms=(700, 3000),
        key_interval_sigma=0.25,
        correction_delay_ms=(150, 1100),
        fatigue_rate=0.0035,
        fatigue_error_multiplier=0.50,
        recovery_rate=0.17,
        uncorrected_probability=0.015,
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
# EMPIRICAL CALIBRATION
# ============================================================

class EmpiricalCalibrator:
    def __init__(
        self,
        path: Path,
        max_rows: int = CALIBRATION_MAX_ROWS,
        trim_fraction: float = CALIBRATION_TRIM_FRACTION,
    ):
        self.path = Path(path)
        self.max_rows = max_rows
        self.trim_fraction = max(0.0, min(0.2, trim_fraction))

    @staticmethod
    def _trim(values: list[float], fraction: float) -> list[float]:
        if not values:
            return []
        values = sorted(values)
        n = int(len(values) * fraction)
        if n * 2 >= len(values):
            return values
        return values[n:len(values) - n]

    @staticmethod
    def _stats(values: list[float]) -> dict[str, Optional[float]]:
        if not values:
            return {
                "count": 0, "mean": None, "median": None,
                "std": None, "p05": None, "p95": None,
            }
        ordered = sorted(values)
        return {
            "count": len(values),
            "mean": statistics.mean(values),
            "median": statistics.median(values),
            "std": statistics.stdev(values) if len(values) > 1 else 0.0,
            "p05": ordered[int(0.05 * (len(ordered) - 1))],
            "p95": ordered[int(0.95 * (len(ordered) - 1))],
        }

    def fit(self) -> dict[str, Any]:
        if not self.path.exists():
            raise FileNotFoundError(f"Reference dataset not found: {self.path}")

        dd_values: list[float] = []
        with self.path.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            for idx, row in enumerate(reader):
                if idx >= self.max_rows:
                    break
                for k, v in row.items():
                    if k and k.startswith("DD."):
                        try:
                            val = float(v)
                            if val > 0:
                                dd_values.append(val)
                        except ValueError:
                            pass

        dd_values = self._trim(dd_values, self.trim_fraction)
        iki = self._stats(dd_values)

        log_sigma = None
        if len(dd_values) > 1:
            log_sigma = statistics.stdev([math.log(x) for x in dd_values])

        return {
            "dataset_type": "CMU_keystroke_dynamics",
            "rows_used": min(self.max_rows, len(dd_values)),
            "inter_key_dd_seconds": iki,
            "lognormal_iki_sigma": log_sigma,
        }

    def apply_to_profile(self, profile: HumanProfile, report: dict[str, Any]) -> HumanProfile:
        import copy
        calibrated = copy.deepcopy(profile)
        iki = report.get("inter_key_dd_seconds") or {}
        iki_mean = iki.get("mean")
        iki_median = iki.get("median")
        iki_p95 = iki.get("p95")

        if iki_mean and iki_mean > 0:
            calibrated.base_wpm = max(
                calibrated.min_wpm,
                min(calibrated.max_wpm, 60.0 / (iki_mean * 5.0)),
            )

        sigma = report.get("lognormal_iki_sigma")
        if sigma is not None:
            calibrated.key_interval_sigma = max(0.08, min(0.70, float(sigma)))

        if iki_median:
            calibrated.word_pause_ms = (
                max(30.0, min(100.0, iki_median * 1000 * 0.45)),
                max(120.0, min(450.0, iki_median * 1000 * 1.5)),
            )

        if iki_p95:
            calibrated.thinking_pause_ms = (
                max(350.0, iki_p95 * 1000),
                max(1200.0, iki_p95 * 1000 * 5.0),
            )

        return calibrated

    def write_report(self, report: dict[str, Any], path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(report, indent=2), encoding="utf-8")


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
            "mean_key_latency_ms": (
                round(statistics.mean(latencies), 3) if latencies else None
            ),
            "median_key_latency_ms": (
                round(statistics.median(latencies), 3) if latencies else None
            ),
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
        logger.info(f"Checkpoint saved: progress up to cell {cell_index}")

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
        if enabled and self.keyboard is None:
            raise RuntimeError("Keyboard execution requested, but pynput is unavailable.")

    def tap(self, key: Any, hold: float = 0.05) -> None:
        if not self.enabled:
            return
        self.keyboard.press(key)
        time.sleep(max(0.005, hold))
        self.keyboard.release(key)

    def backspace(self) -> None:
        if self.enabled:
            self.tap(Key.backspace)

    def tab(self) -> None:
        """Navigates to the next column in Excel."""
        if self.enabled:
            self.tap(Key.tab)

    def enter(self) -> None:
        """Commits cell and returns to start of next row."""
        if self.enabled:
            self.tap(Key.enter)

    def alt_enter(self) -> None:
        """Enters newline within a single Excel cell."""
        if not self.enabled:
            return
        self.keyboard.press(Key.alt)
        time.sleep(0.02)
        self.tap(Key.enter)
        self.keyboard.release(Key.alt)


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

        if error_type == "omission":
            self.emit(
                "key",
                intended=char,
                actual=None,
                error_type="omission",
                row_idx=cell.row_idx,
                col_idx=cell.col_idx,
            )
            if mode == "immediate":
                delay = self.rng.uniform(*self.profile.correction_delay_ms)
                self.sleep_simulated(delay / 1000)
                self.executor.tap(char)
                self.emit(
                    "correction",
                    intended=char,
                    actual=char,
                    corrected=True,
                    row_idx=cell.row_idx,
                    col_idx=cell.col_idx,
                )
                return char, True
            return "", True

        if error_type == "duplication":
            self.executor.tap(char)
            self.executor.tap(char)
            self.emit(
                "key",
                intended=char,
                actual=char + char,
                error_type="duplication",
                row_idx=cell.row_idx,
                col_idx=cell.col_idx,
            )
            self.sleep_simulated(self.behavior.character_delay(None, char) * 2)

            if mode in ("immediate", "delayed"):
                delay = self.rng.uniform(*self.profile.correction_delay_ms)
                self.sleep_simulated(delay / 1000)
                self.executor.backspace()
                self.emit(
                    "correction",
                    intended=char,
                    actual=char,
                    corrected=True,
                    row_idx=cell.row_idx,
                    col_idx=cell.col_idx,
                )
            return char, True

        return char, True

    def type_cell(self, cell: CellItem) -> None:
        text = cell.text
        i = 0
        previous: Optional[str] = None

        while i < len(text):
            char = text[i]

            # In Excel, newline inside cell requires Alt+Enter
            if char == "\n":
                self.executor.alt_enter()
                self.emit(
                    "format",
                    intended="\n",
                    actual="\n",
                    row_idx=cell.row_idx,
                    col_idx=cell.col_idx,
                    metadata={"action": "excel_cell_alt_enter"},
                )
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

            if char in " \t" and self.behavior.should_pause(char):
                self.emit_pause(
                    self.behavior.pause_duration("word_boundary"),
                    "word_boundary",
                    row_idx=cell.row_idx,
                    col_idx=cell.col_idx,
                )

            previous = char
            i += 1

    def run(self, start_index: int = 0) -> dict[str, Any]:
        session_started = datetime.now(timezone.utc)
        sprint_duration = self.rng.uniform(*WORK_SPRINT_MINUTES)
        sprint_start_ms = self.elapsed_ms

        total_cells = len(self.cells)
        logger.info(f"Session started with {total_cells} total cells across rows.")

        for index in range(start_index, total_cells):
            if self.elapsed_ms / 3_600_000 >= MAX_SESSION_HOURS:
                logger.warning("Max session hours reached. Halting.")
                break

            cell = self.cells[index]

            # Work sprint checks
            sprint_elapsed = (self.elapsed_ms - sprint_start_ms) / 1000
            if sprint_elapsed >= sprint_duration * 60:
                duration = self.rng.uniform(*LONG_BREAK_MINUTES)
                logger.info(f"[LONG BREAK] Resting for {duration:.1f} minutes...")
                self.emit("break", metadata={"type": "long", "duration": duration})
                self.sleep_simulated(duration * 60)
                self.behavior.recover(duration)
                sprint_duration = self.rng.uniform(*WORK_SPRINT_MINUTES)
                sprint_start_ms = self.elapsed_ms

            # Type cell contents
            self.type_cell(cell)

            # Move to next cell or next row
            if cell.is_last_in_row:
                # End of row: Enter moves down and aligns to col 1
                self.executor.enter()
                self.emit(
                    "navigation",
                    metadata={"action": "enter_new_row"},
                    row_idx=cell.row_idx,
                    col_idx=cell.col_idx,
                )
                self.rows_since_break += 1
                pause_dur = self.behavior.pause_duration("row_transition")
                self.emit_pause(
                    pause_dur,
                    "row_transition",
                    row_idx=cell.row_idx,
                    col_idx=cell.col_idx,
                )

                # Check micro break after rows
                if self.rows_since_break >= self.next_break_target:
                    brk_duration = self.rng.uniform(*SHORT_BREAK_MINUTES)
                    logger.info(f"[SHORT BREAK] Resting for {brk_duration:.1f} min...")
                    self.sleep_simulated(brk_duration * 60)
                    self.behavior.recover(brk_duration)
                    self.rows_since_break = 0
                    self.next_break_target = self.rng.randint(*SHORT_BREAK_INTERVAL_ROWS)
            else:
                # Next column in same row: Tab
                self.executor.tab()
                self.emit(
                    "navigation",
                    metadata={"action": "tab_next_column"},
                    row_idx=cell.row_idx,
                    col_idx=cell.col_idx,
                )
                pause_dur = self.behavior.pause_duration("cell_transition")
                self.emit_pause(
                    pause_dur,
                    "cell_transition",
                    row_idx=cell.row_idx,
                    col_idx=cell.col_idx,
                )

            # Periodic checkpointing
            if (index + 1) % 25 == 0:
                self.state.save(index + 1, self.session_id, self.seed)

        self.state.save(total_cells, self.session_id, self.seed)
        summary = self.metrics.summary()
        summary.update({
            "session_id": self.session_id,
            "seed": self.seed,
            "profile": self.profile.name,
            "started_at": session_started.isoformat(),
            "finished_at": datetime.now(timezone.utc).isoformat(),
            "csv": str(CSV_PATH),
        })

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
    if not cells:
        raise RuntimeError("No readable rows found in the CSV.")

    logger.info(f"Target CSV: {args.csv} ({len(cells)} cells)")
    logger.info(f"Starting at cell index {start_index}. Focus target Excel window!")

    if config.execute:
        for remaining in range(STARTUP_DELAY_SECONDS, 0, -1):
            print(f"Focus Excel cell. Starting in {remaining}s...", end="\r", flush=True)
            time.sleep(1)
        print(" " * 45, end="\r")

        if human_click is not None:
            human_click(SAFE_EXCEL_CELL_CLICK[0], SAFE_EXCEL_CELL_CLICK[1])

    simulation = ExcelTypingSimulation(
        cells=cells,
        config=config,
        profile=profile,
        seed=seed,
        session_id=session_id,
    )

    summary = simulation.run(start_index=start_index)
    logger.info("Simulation completed.")
    logger.info(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()