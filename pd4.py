"""
AI Training Typing/Data Simulation Engine
-----------------------------------------

A single-file, modular simulator that:
- extracts text/images/basic formatting from a PDF
- builds a document stream
- simulates configurable human typing behavior
- models timing, fatigue, pauses, and multiple error types
- produces a reproducible structured event log
- maintains resumable state
- optionally executes the generated actions through pynput/Windows
- calculates session-quality metrics

Install:
    pip install pymupdf pillow pynput pywin32

Optional:
    The execution adapter expects the same hm.py helpers used by the
    original program: human_click, human_reading_wander, human_scroll.

IMPORTANT:
The simulator is the primary component. The Windows executor is only
an output adapter; the generated event stream can be used independently
as synthetic training data.
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import random
import statistics
import sys
import time
import uuid
from dataclasses import asdict, dataclass, field
from datetime import datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path
from typing import Any, Iterable, Optional

import pymupdf
from PIL import Image

# Optional Windows/execution dependencies.
try:
    from pynput.keyboard import Controller as KeyboardController, Key
except Exception:
    KeyboardController = None
    Key = None

try:
    import win32clipboard
except Exception:
    win32clipboard = None

try:
    from hm import human_click, human_reading_wander, human_scroll
except Exception:
    human_click = human_reading_wander = human_scroll = None


# ============================================================
# CONFIGURATION
# ============================================================

SCRIPT_DIR = Path(__file__).resolve().parent
PDF_PATH = SCRIPT_DIR / "input5.pdf"

OUTPUT_DIR = Path(r"C:\typing_sim\simulation_output")
STATE_FILE = OUTPUT_DIR / "typing_state.json"
EVENT_LOG_FILE = OUTPUT_DIR / "events.jsonl"
SESSION_SUMMARY_FILE = OUTPUT_DIR / "session_summary.json"
CALIBRATION_REPORT_FILE = OUTPUT_DIR / "calibration_report.json"

# Optional local empirical reference dataset. The program will NOT bundle
# third-party research data; point this at a dataset you are permitted to use.
# CMU benchmark CSV:
# https://www.cs.cmu.edu/~keystroke/DSL-StrongPasswordData.csv
REFERENCE_DATA_PATH: Optional[Path] = None
AUTO_DOWNLOAD_CMU_REFERENCE = False
CMU_REFERENCE_URL = "https://www.cs.cmu.edu/~keystroke/DSL-StrongPasswordData.csv"

# When enabled, fit the simulator's timing/error priors from the reference
# data before generating the session.
ENABLE_EMPIRICAL_CALIBRATION = True
CALIBRATION_MAX_ROWS = 250_000
CALIBRATION_TRIM_FRACTION = 0.01

# Set to False for a pure offline simulation that does not touch
# the keyboard/mouse.
EXECUTE_IN_APPLICATION = True

# Reproducibility:
# Set to an integer for a reproducible session.
# Set to None for a fresh random seed every run.
RANDOM_SEED: Optional[int] = 20261004

# Choose a human profile.
PROFILE_NAME = "average"

# Work schedule.
STARTUP_DELAY_SECONDS = 10
WORK_SPRINT_MINUTES = (75, 105)
LONG_BREAK_MINUTES = (12, 18)
SHORT_BREAK_MINUTES = (1, 3)
SHORT_BREAK_INTERVAL_LINES = (35, 60)

# Application coordinates retained from the original program.
SAFE_TITLE_BAR_CLICK = (600, 15)

# Safety: do not allow an accidental runaway execution.
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
    punctuation_pause_ms: tuple[float, float]
    thinking_pause_ms: tuple[float, float]

    key_interval_sigma: float
    correction_delay_ms: tuple[float, float]

    fatigue_rate: float
    fatigue_error_multiplier: float
    recovery_rate: float

    # Probability that a generated error is intentionally left
    # uncorrected.
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
        punctuation_pause_ms=(100, 350),
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
        punctuation_pause_ms=(120, 500),
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
        punctuation_pause_ms=(180, 650),
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
        if self.seed is not None and not isinstance(self.seed, int):
            raise TypeError("seed must be an integer or None")


@dataclass
class DocumentItem:
    item_id: int
    item_type: str
    text: str = ""
    is_bold: bool = False
    is_italic: bool = False
    is_underline: bool = False
    page: int = 0
    bbox: Optional[tuple[float, float, float, float]] = None
    image_bytes: Optional[bytes] = None


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

    item_id: Optional[int] = None
    page: Optional[int] = None
    wpm: Optional[float] = None
    fatigue: Optional[float] = None
    formatting: Optional[dict[str, bool]] = None
    metadata: dict[str, Any] = field(default_factory=dict)


# ============================================================
# KEYBOARD MODEL
# ============================================================

# Approximate QWERTY coordinates. Coordinates are intentionally
# continuous so physical distance can influence substitution odds.
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

# Retained as a fallback and for modifier-independent error selection.
SHIFTED = {
    "!": "1", "@": "2", "#": "3", "$": "4", "%": "5",
    "^": "6", "&": "7", "*": "8", "(": "9", ")": "0",
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
    return [
        k for k in KEY_POSITIONS
        if k != c and key_distance(c, k) <= radius
    ]


# ============================================================
# PDF/DOCUMENT PARSER
# ============================================================

class PDFDocumentParser:
    """Extracts text/images and retains basic formatting/geometry."""

    def __init__(self, pdf_path: Path):
        self.pdf_path = pdf_path

    @staticmethod
    def _span_flags(span: dict[str, Any]) -> tuple[bool, bool, bool]:
        flags = span.get("flags", 0)
        font = span.get("font", "").lower()

        bold = bool(flags & 16) or any(
            x in font for x in ("bold", "heavy", "black")
        )

        italic = bool(flags & 2) or any(
            x in font for x in ("italic", "oblique")
        )

        # PyMuPDF does not expose underline uniformly through span flags,
        # so this remains conservative.
        underline = bool(flags & 4)

        return bold, italic, underline

    def parse(self) -> list[DocumentItem]:
        if not self.pdf_path.exists():
            raise FileNotFoundError(self.pdf_path)

        doc = pymupdf.open(self.pdf_path)
        items: list[DocumentItem] = []
        next_id = 0

        try:
            for page_number, page in enumerate(doc, start=1):
                blocks = page.get_text("dict").get("blocks", [])

                # Sort by vertical position first, then horizontal position.
                # This is substantially safer than blindly trusting source order.
                blocks = sorted(
                    blocks,
                    key=lambda b: (
                        round(float(b.get("bbox", [0, 0, 0, 0])[1]), 1),
                        round(float(b.get("bbox", [0, 0, 0, 0])[0]), 1),
                    ),
                )

                for block in blocks:
                    block_type = block.get("type")
                    bbox = tuple(block.get("bbox", (0, 0, 0, 0)))

                    if block_type == 0:
                        lines = block.get("lines", [])

                        for line in lines:
                            spans = line.get("spans", [])
                            for span in spans:
                                text = span.get("text", "")
                                if not text:
                                    continue

                                bold, italic, underline = self._span_flags(span)

                                items.append(
                                    DocumentItem(
                                        item_id=next_id,
                                        item_type="text",
                                        text=text,
                                        is_bold=bold,
                                        is_italic=italic,
                                        is_underline=underline,
                                        page=page_number,
                                        bbox=bbox,
                                    )
                                )
                                next_id += 1

                            # Preserve visual line boundaries.
                            items.append(
                                DocumentItem(
                                    item_id=next_id,
                                    item_type="soft_break",
                                    text="\n",
                                    page=page_number,
                                )
                            )
                            next_id += 1

                    elif block_type == 1:
                        image_bytes = block.get("image")
                        if image_bytes:
                            items.append(
                                DocumentItem(
                                    item_id=next_id,
                                    item_type="image",
                                    image_bytes=image_bytes,
                                    page=page_number,
                                    bbox=bbox,
                                )
                            )
                            next_id += 1

                if page_number < len(doc):
                    items.append(
                        DocumentItem(
                            item_id=next_id,
                            item_type="paragraph_break",
                            text="\n\n",
                            page=page_number,
                        )
                    )
                    next_id += 1
        finally:
            doc.close()

        return items


# ==========================================================
# EMPIRICAL CALIBRATION
# ============================================================

class EmpiricalCalibrator:
    """Fit simulator priors from real keystroke reference data.

    Supported reference styles:
      1. CMU DSL-StrongPasswordData.csv, which contains H.*, DD.* and UD.*
         timing features in seconds.
      2. Generic event CSV/TSV files containing key/timestamp information.

    The calibrator deliberately does not manufacture measurements that the
    reference data does not contain. For example, CMU's fixed-password
    benchmark is excellent for hold/digraph timing but is not a general
    free-text WPM dataset, so its WPM estimate is labelled as an equivalent
    rate rather than treated as a literal typing-speed measurement.
    """

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
    def _finite(values: Iterable[float]) -> list[float]:
        out = []
        for value in values:
            try:
                x = float(value)
            except (TypeError, ValueError):
                continue
            if math.isfinite(x):
                out.append(x)
        return out

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
                "count": 0,
                "mean": None,
                "median": None,
                "std": None,
                "p05": None,
                "p95": None,
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

    @staticmethod
    def _read_rows(path: Path) -> Iterable[dict[str, str]]:
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            sample = f.read(8192)
            f.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\t| ")
            except csv.Error:
                dialect = csv.excel
                dialect.delimiter = ","

            reader = csv.DictReader(f, dialect=dialect)
            for row in reader:
                yield {str(k).strip(): (v or "").strip() for k, v in row.items() if k is not None}

    def _looks_like_cmu(self, fieldnames: Iterable[str]) -> bool:
        names = list(fieldnames)
        return any(x.startswith("H.") for x in names) and any(
            x.startswith("DD.") for x in names
        )

    def _calibrate_cmu(self, rows: list[dict[str, str]]) -> dict[str, Any]:
        h_values = []
        dd_values = []
        ud_values = []
        subjects = set()

        for row in rows:
            if row.get("subject"):
                subjects.add(row["subject"])

            for key, value in row.items():
                if key.startswith("H."):
                    try:
                        h_values.append(float(value))
                    except ValueError:
                        pass
                elif key.startswith("DD."):
                    try:
                        dd_values.append(float(value))
                    except ValueError:
                        pass
                elif key.startswith("UD."):
                    try:
                        ud_values.append(float(value))
                    except ValueError:
                        pass

        h_values = self._trim(
            [x for x in h_values if x > 0], self.trim_fraction
        )
        dd_values = self._trim(
            [x for x in dd_values if x > 0], self.trim_fraction
        )

        # DD is a down-down digraph interval, so it is the closest direct
        # reference for the simulator's inter-key interval.
        iki = self._stats(dd_values)
        hold = self._stats(h_values)

        positive_ud = [x for x in ud_values if x > 0]
        negative_ud = [x for x in ud_values if x < 0]

        # Convert the empirical IKI distribution to the lognormal sigma used
        # by the simulator's character-delay generator.
        log_sigma = None
        if len(dd_values) > 1:
            logs = [math.log(x) for x in dd_values]
            log_sigma = statistics.stdev(logs)

        # Fixed-password data has no natural word boundaries. This is an
        # equivalent 5-character WPM based solely on the observed DD mean.
        equivalent_wpm = None
        if iki["mean"] and iki["mean"] > 0:
            equivalent_wpm = 60.0 / (iki["mean"] * 5.0)

        return {
            "dataset_type": "CMU_keystroke_dynamics",
            "subjects": len(subjects),
            "rows_used": len(rows),
            "hold_seconds": hold,
            "inter_key_dd_seconds": iki,
            "keyup_keydown_seconds": self._stats(ud_values),
            "positive_ud_fraction": (
                len(positive_ud) / len(ud_values) if ud_values else None
            ),
            "negative_ud_fraction": (
                len(negative_ud) / len(ud_values) if ud_values else None
            ),
            "lognormal_iki_sigma": log_sigma,
            "equivalent_wpm_from_dd": equivalent_wpm,
            "limitations": [
                "CMU benchmark uses a fixed password rather than natural free text.",
                "Use its timing features primarily for key timing calibration.",
                "Do not interpret equivalent_wpm as a population free-text WPM measurement.",
            ],
        }

    def _calibrate_generic(self, rows: list[dict[str, str]]) -> dict[str, Any]:
        if not rows:
            return {"dataset_type": "generic", "rows_used": 0}

        names = list(rows[0].keys())
        lower = {n.lower(): n for n in names}

        def find(*candidates: str) -> Optional[str]:
            for candidate in candidates:
                if candidate in lower:
                    return lower[candidate]
            return None

        key_col = find("key", "key_pressed", "keyname", "key_name", "symbol")
        timestamp_col = find("timestamp", "time", "time_ms", "timestamp_ms", "ts")
        event_col = find("event", "event_type", "type")
        subject_col = find("subject", "participant", "user", "participant_id")
        wpm_col = find("wpm", "words_per_minute")
        error_col = find("error", "is_error", "error_type")

        timestamps = []
        wpm = []
        errors = []
        subjects = set()
        previous = None
        iki = []

        for row in rows:
            if subject_col and row.get(subject_col):
                subjects.add(row[subject_col])

            if wpm_col:
                try:
                    value = float(row[wpm_col])
                    if value > 0:
                        wpm.append(value)
                except ValueError:
                    pass

            if error_col:
                value = row.get(error_col, "").lower()
                if value in {"1", "true", "yes", "error", "incorrect"} or (
                    value and value not in {"0", "false", "no", "correct"}
                    and "error" in error_col.lower()
                ):
                    errors.append(1)

            if timestamp_col:
                try:
                    t = float(row[timestamp_col])
                    # Normalize obvious millisecond timestamps.
                    if abs(t) > 1e11:
                        t /= 1000.0
                    timestamps.append(t)
                    if previous is not None:
                        delta = t - previous
                        if 0.0 < delta < 5.0:
                            iki.append(delta)
                    previous = t
                except ValueError:
                    pass

        iki = self._trim(iki, self.trim_fraction)

        result = {
            "dataset_type": "generic_event_data",
            "subjects": len(subjects),
            "rows_used": len(rows),
            "columns_detected": {
                "key": key_col,
                "timestamp": timestamp_col,
                "event": event_col,
                "subject": subject_col,
                "wpm": wpm_col,
                "error": error_col,
            },
            "inter_key_seconds": self._stats(iki),
            "wpm": self._stats(wpm),
            "error_rate": (
                sum(errors) / len(rows) if error_col else None
            ),
        }

        return result

    def fit(self) -> dict[str, Any]:
        if not self.path.exists():
            raise FileNotFoundError(
                f"Reference dataset not found: {self.path}"
            )

        rows = []
        for index, row in enumerate(self._read_rows(self.path)):
            rows.append(row)
            if index + 1 >= self.max_rows:
                break

        if not rows:
            raise ValueError("Reference dataset contains no readable rows.")

        if self._looks_like_cmu(rows[0].keys()):
            return self._calibrate_cmu(rows)

        return self._calibrate_generic(rows)

    def apply_to_profile(
        self,
        profile: HumanProfile,
        report: dict[str, Any],
    ) -> HumanProfile:
        """Return a calibrated copy of a profile without mutating the preset."""
        import copy
        calibrated = copy.deepcopy(profile)

        iki = report.get("inter_key_dd_seconds") or report.get("inter_key_seconds")
        wpm = report.get("wpm")
        error_rate = report.get("error_rate")

        # Timing calibration: prefer a direct IKI median/mean.
        iki_mean = iki.get("mean") if isinstance(iki, dict) else None
        iki_p95 = iki.get("p95") if isinstance(iki, dict) else None
        iki_median = iki.get("median") if isinstance(iki, dict) else None

        if iki_mean and iki_mean > 0:
            # WPM-equivalent based on five characters per word.
            empirical_wpm = 60.0 / (iki_mean * 5.0)
            calibrated.base_wpm = max(
                calibrated.min_wpm,
                min(calibrated.max_wpm, empirical_wpm),
            )

        if isinstance(wpm, dict) and wpm.get("median"):
            empirical_wpm = float(wpm["median"])
            calibrated.base_wpm = max(
                calibrated.min_wpm,
                min(calibrated.max_wpm, empirical_wpm),
            )

        # Fit the timing noise from the observed positive IKI distribution.
        sigma = report.get("lognormal_iki_sigma")
        if sigma is None and iki_mean and iki_median and iki_mean > 0:
            # Fallback approximation for positive right-skewed distributions.
            ratio = max(1.0001, iki_mean / iki_median)
            sigma = min(0.65, math.sqrt(2.0 * math.log(ratio)))

        if sigma is not None:
            calibrated.key_interval_sigma = max(
                0.08,
                min(0.70, float(sigma)),
            )

        # Error calibration is only applied when the reference actually has
        # error labels. CMU timing data does not contain natural free-text
        # error labels, so its error rate is intentionally not invented.
        if error_rate is not None and 0 < error_rate < 0.30:
            calibrated.base_error_rate = max(
                0.002,
                min(0.08, float(error_rate)),
            )

        # Use observed timing range to improve realistic pause-free key timing.
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
        path.write_text(
            json.dumps(report, indent=2),
            encoding="utf-8",
        )


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
        value = self.rng.gauss(
            self.profile.base_wpm,
            self.profile.wpm_stddev,
        )
        return max(
            self.profile.min_wpm,
            min(self.profile.max_wpm, value),
        )

    def recover(self, minutes: float) -> None:
        self.fatigue *= max(
            0.0,
            math.exp(-self.profile.recovery_rate * minutes),
        )
        self.current_wpm = min(
            self.profile.max_wpm,
            self.current_wpm + self.profile.wpm_stddev * 0.25,
        )

    def advance(self, elapsed_minutes: float) -> None:
        self.fatigue = min(
            1.0,
            self.fatigue + elapsed_minutes * self.profile.fatigue_rate,
        )

        drift = self.rng.gauss(0, self.profile.wpm_stddev * 0.025)

        fatigue_factor = max(0.60, 1.0 - self.fatigue * 0.25)

        target = self.profile.base_wpm * fatigue_factor

        self.current_wpm += (
            (target - self.current_wpm) * 0.03
            + drift
        )

        self.current_wpm = max(
            self.profile.min_wpm,
            min(self.profile.max_wpm, self.current_wpm),
        )

    def error_probability(self, char: str) -> float:
        punctuation_multiplier = 1.5 if char in ".,;:!?()[]{}'\"-" else 1.0
        fatigue_multiplier = 1.0 + (
            self.fatigue * self.profile.fatigue_error_multiplier
        )

        return min(
            0.25,
            self.profile.base_error_rate
            * punctuation_multiplier
            * fatigue_multiplier,
        )

    def base_char_delay(self) -> float:
        # WPM -> approximate seconds/character.
        cps = max(1.0, self.current_wpm * 5.0 / 60.0)
        return 1.0 / cps

    def character_delay(self, previous: Optional[str], current: str) -> float:
        base = self.base_char_delay()

        multiplier = 1.0

        if previous and previous.lower() in FINGER and current.lower() in FINGER:
            if FINGER[previous.lower()] == FINGER[current.lower()]:
                multiplier += 0.10
            elif FINGER[previous.lower()].startswith(
                FINGER[current.lower()][0]
            ):
                multiplier += 0.025
            else:
                multiplier -= 0.025

        if current in ".,;:!?":
            multiplier += 0.10

        # Log-normal noise creates positive timing values and a long tail.
        delay = self.rng.lognormvariate(
            math.log(max(0.025, base * multiplier)),
            self.profile.key_interval_sigma,
        )

        return max(0.025, min(1.8, delay))

    def should_pause(self, current: str) -> bool:
        probability = self.profile.pause_probability

        if current in ".!?":
            probability *= 2.0

        return self.rng.random() < probability

    def pause_duration(self, reason: str) -> float:
        if reason == "word_boundary":
            lo, hi = self.profile.word_pause_ms
        elif reason == "punctuation":
            lo, hi = self.profile.punctuation_pause_ms
        else:
            lo, hi = self.profile.thinking_pause_ms

        # Log-normal-ish positive pause distribution.
        midpoint = max(1.0, (lo + hi) / 2)
        sigma = 0.45
        value = self.rng.lognormvariate(math.log(midpoint), sigma)
        return max(lo, min(hi * 2.5, value))

    def choose_error(self, char: str) -> Optional[str]:
        p = self.error_probability(char)

        if self.rng.random() >= p:
            return None

        choices: list[tuple[str, float]] = []

        if char.isalpha():
            choices.extend([
                ("substitution", 0.45),
                ("omission", self.profile.omission_probability),
                ("duplication", self.profile.duplication_probability),
                ("transposition", self.profile.transposition_probability),
            ])

        if char in ".,;:!?":
            choices.append(
                ("punctuation", self.profile.punctuation_error_probability)
            )

        if char.isalpha():
            choices.append(
                ("capitalization", self.profile.capitalization_error_probability)
            )

        if char == " ":
            choices.append(
                ("space", self.profile.space_error_probability)
            )

        names = [name for name, weight in choices if weight > 0]
        weights = [weight for _, weight in choices if weight > 0]

        if not names:
            return "substitution"

        return self.rng.choices(names, weights=weights, k=1)[0]

    def correction_mode(self) -> str:
        r = self.rng.random()

        if r < self.profile.uncorrected_probability:
            return "none"

        if r < self.profile.uncorrected_probability + self.profile.delayed_correction_probability:
            return "delayed"

        if r < self.profile.uncorrected_probability + self.profile.delayed_correction_probability + self.profile.correction_probability:
            return "immediate"

        return "none"


# ============================================================
# EVENT LOGGER
# ============================================================

class EventLogger:
    def __init__(self, path: Path):
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.event_count = 0

        # Start a fresh event log for a new session.
        self.path.write_text("", encoding="utf-8")

    def write(self, event: Event) -> None:
        self.event_count += 1
        payload = asdict(event)
        with self.path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(payload, ensure_ascii=False) + "\n")


# ============================================================
# METRICS
# ============================================================

class Metrics:
    def __init__(self):
        self.events: list[Event] = []

    def add(self, event: Event) -> None:
        self.events.append(event)

    def summary(self) -> dict[str, Any]:
        if not self.events:
            return {"event_count": 0}

        key_events = [
            e for e in self.events
            if e.event_type == "key"
        ]

        errors = [
            e for e in self.events
            if e.error_type
        ]

        corrections = [
            e for e in self.events
            if e.event_type == "correction"
        ]

        pauses = [
            e for e in self.events
            if e.event_type == "pause"
        ]

        elapsed_ms = max(e.elapsed_ms for e in self.events)

        actual_chars = sum(
            1 for e in key_events
            if e.actual is not None
        )

        words = sum(
            1 for e in key_events
            if e.actual == " "
        )

        minutes = max(0.001, elapsed_ms / 60000)

        wpm = words / minutes

        latencies = [
            e.latency_ms
            for e in key_events
            if e.latency_ms is not None
        ]

        error_rate = (
            len(errors) / max(1, len(key_events))
        )

        correction_rate = (
            len(corrections) / max(1, len(errors))
        )

        pause_durations = [
            e.metadata.get("duration_ms", 0)
            for e in pauses
        ]

        flags = []

        if wpm > 150:
            flags.append("unusually_high_wpm")

        if wpm < 15 and actual_chars > 200:
            flags.append("unusually_low_wpm")

        if actual_chars > 2000 and len(errors) == 0:
            flags.append("zero_error_long_session")

        if corrections and correction_rate > 1.05:
            flags.append("correction_count_anomaly")

        if latencies:
            mean_latency = statistics.mean(latencies)
            median_latency = statistics.median(latencies)
        else:
            mean_latency = median_latency = None

        return {
            "event_count": len(self.events),
            "key_events": len(key_events),
            "characters": actual_chars,
            "spaces": words,
            "duration_minutes": round(minutes, 3),
            "effective_wpm": round(wpm, 3),
            "error_count": len(errors),
            "error_rate": round(error_rate, 6),
            "correction_count": len(corrections),
            "correction_rate_per_error": round(correction_rate, 4),
            "pause_count": len(pauses),
            "mean_key_latency_ms": (
                round(mean_latency * 1, 3)
                if mean_latency is not None else None
            ),
            "median_key_latency_ms": (
                round(median_latency * 1, 3)
                if median_latency is not None else None
            ),
            "mean_pause_ms": (
                round(statistics.mean(pause_durations), 3)
                if pause_durations else None
            ),
            "quality_flags": flags,
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
            return {
                "item_index": 0,
                "session_id": None,
                "seed": None,
            }

        try:
            return json.loads(self.path.read_text(encoding="utf-8"))
        except Exception as exc:
            print(f"Warning: state file could not be read: {exc}")
            return {
                "item_index": 0,
                "session_id": None,
                "seed": None,
            }

    def save(self, item_index: int, session_id: str, seed: int) -> None:
        payload = {
            "item_index": item_index,
            "session_id": session_id,
            "seed": seed,
            "updated_at": datetime.now(timezone.utc).isoformat(),
        }

        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(
            json.dumps(payload, indent=2),
            encoding="utf-8",
        )
        tmp.replace(self.path)

    def clear(self) -> None:
        if self.path.exists():
            self.path.unlink()


# ============================================================
# WINDOWS EXECUTION ADAPTER
# ============================================================

class WindowsExecutor:
    def __init__(self, enabled: bool):
        self.enabled = enabled
        self.keyboard = (
            KeyboardController()
            if enabled and KeyboardController is not None
            else None
        )

        if enabled and self.keyboard is None:
            raise RuntimeError(
                "Keyboard execution requested, but pynput is unavailable."
            )

    @staticmethod
    def _copy_image_to_clipboard(image_bytes: bytes) -> None:
        if win32clipboard is None:
            raise RuntimeError("pywin32 is required for image clipboard operations.")

        image = Image.open(BytesIO(image_bytes))
        output = BytesIO()
        image.convert("RGB").save(output, "BMP")
        data = output.getvalue()[14:]
        output.close()

        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(
                win32clipboard.CF_DIB,
                data,
            )
        finally:
            win32clipboard.CloseClipboard()

    def tap(self, key: Any, hold: float = 0.06) -> None:
        if not self.enabled:
            return

        self.keyboard.press(key)
        time.sleep(max(0.005, hold))
        self.keyboard.release(key)

    def text(self, text: str) -> None:
        if not self.enabled:
            return

        for char in text:
            self.tap(char)

    def toggle_bold(self) -> None:
        if not self.enabled:
            return

        self.keyboard.press(Key.ctrl)
        time.sleep(0.05)
        self.tap("b", hold=0.06)
        self.keyboard.release(Key.ctrl)
        time.sleep(0.15)

    def backspace(self) -> None:
        if self.enabled:
            self.tap(Key.backspace)

    def enter(self, soft: bool = False) -> None:
        if not self.enabled:
            return

        if soft:
            self.keyboard.press(Key.shift)
            self.tap(Key.enter)
            self.keyboard.release(Key.shift)
        else:
            self.tap(Key.enter)

    def paste_image(self, image_bytes: bytes) -> None:
        if not self.enabled:
            return

        self._copy_image_to_clipboard(image_bytes)

        self.keyboard.press(Key.ctrl)
        time.sleep(0.05)
        self.tap("v")
        self.keyboard.release(Key.ctrl)


# ============================================================
# SIMULATION ENGINE
# ============================================================

class TypingSimulation:
    def __init__(
        self,
        items: list[DocumentItem],
        config: SimulationConfig,
        profile: HumanProfile,
        seed: int,
        session_id: str,
    ):
        self.items = items
        self.config = config
        self.profile = profile
        self.seed = seed
        self.session_id = session_id

        self.rng = random.Random(seed)
        self.behavior = HumanBehaviorModel(profile, self.rng)

        self.logger = EventLogger(config.event_log)
        self.metrics = Metrics()
        self.state = StateManager(config.state_file)
        self.executor = WindowsExecutor(config.execute)

        self.event_id = 0
        self.elapsed_ms = 0.0
        self.start_time = datetime.now(timezone.utc)
        self.current_formatting = {
            "bold": False,
            "italic": False,
            "underline": False,
        }

        self.lines_since_break = 0
        self.next_break_target = self.rng.randint(
            SHORT_BREAK_INTERVAL_LINES[0],
            SHORT_BREAK_INTERVAL_LINES[1],
        )

    # --------------------------------------------------------
    # EVENT CREATION
    # --------------------------------------------------------

    def emit(
        self,
        event_type: str,
        *,
        intended: Optional[str] = None,
        actual: Optional[str] = None,
        latency_ms: Optional[float] = None,
        error_type: Optional[str] = None,
        corrected: Optional[bool] = None,
        correction_delay_ms: Optional[float] = None,
        item: Optional[DocumentItem] = None,
        metadata: Optional[dict[str, Any]] = None,
    ) -> Event:
        self.event_id += 1

        event = Event(
            event_id=self.event_id,
            session_id=self.session_id,
            timestamp=datetime.now(timezone.utc).isoformat(),
            elapsed_ms=round(self.elapsed_ms, 3),
            event_type=event_type,
            intended=intended,
            actual=actual,
            latency_ms=latency_ms,
            error_type=error_type,
            corrected=corrected,
            correction_delay_ms=correction_delay_ms,
            item_id=item.item_id if item else None,
            page=item.page if item else None,
            wpm=round(self.behavior.current_wpm, 3),
            fatigue=round(self.behavior.fatigue, 6),
            formatting=dict(self.current_formatting),
            metadata=metadata or {},
        )

        self.logger.write(event)
        self.metrics.add(event)
        return event

    def sleep_simulated(self, seconds: float) -> None:
        self.elapsed_ms += seconds * 1000

        # In offline mode, simulated time advances without real waiting.
        if self.config.execute:
            time.sleep(seconds)

        self.behavior.advance(seconds / 60)

    def emit_pause(self, duration_ms: float, reason: str, item=None) -> None:
        self.emit(
            "pause",
            item=item,
            metadata={
                "duration_ms": round(duration_ms, 3),
                "reason": reason,
            },
        )
        self.sleep_simulated(duration_ms / 1000)

    # --------------------------------------------------------
    # FORMATTING
    # --------------------------------------------------------

    def apply_formatting(self, item: DocumentItem) -> None:
        desired = {
            "bold": item.is_bold,
            "italic": item.is_italic,
            "underline": item.is_underline,
        }

        # The original execution layer only directly supports bold.
        # We still record the other formatting states in the event stream.
        if desired["bold"] != self.current_formatting["bold"]:
            if self.config.execute:
                self.executor.toggle_bold()

            self.current_formatting["bold"] = desired["bold"]

            self.emit(
                "format",
                item=item,
                metadata={
                    "action": "toggle_bold",
                    "enabled": desired["bold"],
                },
            )

        self.current_formatting["italic"] = desired["italic"]
        self.current_formatting["underline"] = desired["underline"]

    # --------------------------------------------------------
    # ERROR GENERATION
    # --------------------------------------------------------

    def substitution(self, char: str) -> Optional[str]:
        candidates = nearby_keys(char)

        if not candidates:
            return None

        # Closer keys have substantially higher probability.
        weights = []
        for candidate in candidates:
            d = key_distance(char, candidate)
            weights.append(math.exp(-1.5 * d))

        return self.rng.choices(
            candidates,
            weights=weights,
            k=1,
        )[0]

    def process_error(
        self,
        char: str,
        error_type: str,
        item: DocumentItem,
    ) -> tuple[str, bool]:
        """
        Returns:
            actual character generated for the current logical position,
            whether the original character should be advanced.
        """

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
                corrected=False,
                item=item,
                metadata={
                    "distance": round(key_distance(char, wrong), 4)
                },
            )

            self.executor.tap(wrong)
            self.sleep_simulated(
                self.behavior.character_delay(None, wrong)
            )

            if mode == "immediate":
                delay = self.rng.uniform(
                    *self.profile.correction_delay_ms
                )

                self.executor.backspace()
                self.sleep_simulated(delay / 1000)

                self.emit(
                    "correction",
                    intended=char,
                    actual=char,
                    corrected=True,
                    correction_delay_ms=delay,
                    item=item,
                    metadata={
                        "correction_type": "backspace_and_retry"
                    },
                )

                self.executor.tap(char)
                self.sleep_simulated(
                    self.behavior.character_delay(wrong, char)
                )

                return char, True

            if mode == "delayed":
                delay = self.rng.uniform(
                    *self.profile.correction_delay_ms
                ) * self.rng.uniform(1.5, 4.0)

                self.sleep_simulated(delay / 1000)

                # Delayed correction is represented in the event stream.
                self.executor.backspace()
                self.executor.tap(char)

                self.emit(
                    "correction",
                    intended=char,
                    actual=char,
                    corrected=True,
                    correction_delay_ms=delay,
                    item=item,
                    metadata={
                        "correction_type": "delayed"
                    },
                )

                self.sleep_simulated(
                    self.behavior.character_delay(wrong, char)
                )

                return char, True

            return wrong, True

        if error_type == "omission":
            self.emit(
                "key",
                intended=char,
                actual=None,
                error_type="omission",
                corrected=(mode != "none"),
                item=item,
            )

            if mode == "immediate":
                # The user realizes the omission and types it after a pause.
                delay = self.rng.uniform(
                    *self.profile.correction_delay_ms
                )

                self.sleep_simulated(delay / 1000)
                self.executor.tap(char)

                self.emit(
                    "correction",
                    intended=char,
                    actual=char,
                    corrected=True,
                    correction_delay_ms=delay,
                    item=item,
                    metadata={
                        "correction_type": "omission_repair"
                    },
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
                corrected=False,
                item=item,
            )

            self.sleep_simulated(
                self.behavior.character_delay(None, char) * 2
            )

            if mode in ("immediate", "delayed"):
                delay = self.rng.uniform(
                    *self.profile.correction_delay_ms
                )

                self.executor.backspace()
                self.sleep_simulated(delay / 1000)

                self.emit(
                    "correction",
                    intended=char,
                    actual=char,
                    corrected=True,
                    correction_delay_ms=delay,
                    item=item,
                    metadata={
                        "correction_type": "duplicate_removal"
                    },
                )

            return char, True

        if error_type == "capitalization":
            wrong = char.lower() if char.isupper() else char.upper()

            self.executor.tap(wrong)
            self.sleep_simulated(
                self.behavior.character_delay(None, wrong)
            )

            self.emit(
                "key",
                intended=char,
                actual=wrong,
                error_type="capitalization",
                corrected=False,
                item=item,
            )

            if mode in ("immediate", "delayed"):
                delay = self.rng.uniform(
                    *self.profile.correction_delay_ms
                )

                self.executor.backspace()
                self.executor.tap(char)
                self.sleep_simulated(delay / 1000)

                self.emit(
                    "correction",
                    intended=char,
                    actual=char,
                    corrected=True,
                    correction_delay_ms=delay,
                    item=item,
                    metadata={
                        "correction_type": "capitalization_repair"
                    },
                )

            return wrong, True

        if error_type == "space":
            # Randomly omit or duplicate spaces.
            if self.rng.random() < 0.5:
                self.emit(
                    "key",
                    intended=" ",
                    actual=None,
                    error_type="space_omission",
                    corrected=False,
                    item=item,
                )
                return "", True

            self.executor.tap(" ")
            self.executor.tap(" ")

            self.emit(
                "key",
                intended=" ",
                actual="  ",
                error_type="space_duplication",
                corrected=False,
                item=item,
            )

            if mode in ("immediate", "delayed"):
                self.executor.backspace()

            return " ", True

        if error_type == "punctuation":
            alternatives = {
                ",": ".",
                ".": ",",
                ";": ":",
                ":": ";",
                "!": "?",
                "?": "!",
            }

            wrong = alternatives.get(
                char,
                self.substitution(char) or char,
            )

            self.executor.tap(wrong)

            self.emit(
                "key",
                intended=char,
                actual=wrong,
                error_type="punctuation",
                corrected=False,
                item=item,
            )

            self.sleep_simulated(
                self.behavior.character_delay(None, wrong)
            )

            if mode in ("immediate", "delayed"):
                delay = self.rng.uniform(
                    *self.profile.correction_delay_ms
                )

                self.executor.backspace()
                self.executor.tap(char)

                self.emit(
                    "correction",
                    intended=char,
                    actual=char,
                    corrected=True,
                    correction_delay_ms=delay,
                    item=item,
                    metadata={
                        "correction_type": "punctuation_repair"
                    },
                )

            return wrong, True

        if error_type == "transposition":
            # Actual transposition is handled in the main text loop,
            # because it requires looking at the next character.
            return char, True

        return char, True

    # --------------------------------------------------------
    # TEXT
    # --------------------------------------------------------

    def type_text(self, item: DocumentItem) -> None:
        text = item.text
        i = 0
        previous: Optional[str] = None

        while i < len(text):
            char = text[i]

            # Preserve whitespace and line handling.
            if char == "\n":
                self.executor.enter(soft=True)
                self.emit(
                    "format",
                    intended="\n",
                    actual="\n",
                    item=item,
                    metadata={"action": "soft_break"},
                )
                self.sleep_simulated(self.rng.uniform(0.25, 0.8))
                self.lines_since_break += 1
                i += 1
                previous = "\n"
                continue

            # Possible transposition.
            if (
                i + 1 < len(text)
                and char.isalpha()
                and text[i + 1].isalpha()
                and self.rng.random()
                < self.behavior.error_probability(char)
                * self.profile.transposition_probability
            ):
                second = text[i + 1]

                self.executor.tap(second)
                self.executor.tap(char)

                self.emit(
                    "key",
                    intended=char + second,
                    actual=second + char,
                    error_type="transposition",
                    corrected=False,
                    item=item,
                )

                delay1 = self.behavior.character_delay(previous, second)
                delay2 = self.behavior.character_delay(second, char)

                self.sleep_simulated(delay1 + delay2)

                mode = self.behavior.correction_mode()

                if mode in ("immediate", "delayed"):
                    delay = self.rng.uniform(
                        *self.profile.correction_delay_ms
                    )

                    # Correcting a two-character transposition.
                    self.executor.backspace()
                    self.executor.backspace()
                    self.executor.tap(char)
                    self.executor.tap(second)

                    self.sleep_simulated(delay / 1000)

                    self.emit(
                        "correction",
                        intended=char + second,
                        actual=char + second,
                        corrected=True,
                        correction_delay_ms=delay,
                        item=item,
                        metadata={
                            "correction_type": "transposition_repair"
                        },
                    )

                i += 2
                previous = second
                continue

            error_type = self.behavior.choose_error(char)

            if error_type:
                actual, advance = self.process_error(
                    char,
                    error_type,
                    item,
                )

                if actual:
                    previous = actual[-1]

                if advance:
                    i += 1

                # Pause occasionally at semantic boundaries.
                if char == " " and self.behavior.should_pause(char):
                    duration = self.behavior.pause_duration(
                        "word_boundary"
                    )
                    self.emit_pause(
                        duration,
                        "word_boundary",
                        item=item,
                    )

                continue

            # Normal key.
            delay = self.behavior.character_delay(
                previous,
                char,
            )

            self.executor.tap(char)

            self.emit(
                "key",
                intended=char,
                actual=char,
                latency_ms=delay * 1000,
                item=item,
            )

            self.sleep_simulated(delay)

            if char in " \t":
                self.lines_since_break += 0

                if self.behavior.should_pause(char):
                    duration = self.behavior.pause_duration(
                        "word_boundary"
                    )
                    self.emit_pause(
                        duration,
                        "word_boundary",
                        item=item,
                    )

            elif char in ".!?":
                if self.behavior.should_pause(char):
                    duration = self.behavior.pause_duration(
                        "punctuation"
                    )
                    self.emit_pause(
                        duration,
                        "punctuation",
                        item=item,
                    )

            elif self.behavior.should_pause(char):
                duration = self.behavior.pause_duration(
                    "thinking"
                )
                self.emit_pause(
                    duration,
                    "thinking",
                    item=item,
                )

            previous = char
            i += 1

    # --------------------------------------------------------
    # IMAGES / BREAKS
    # --------------------------------------------------------

    def insert_image(self, item: DocumentItem) -> None:
        if item.image_bytes:
            self.executor.paste_image(item.image_bytes)

        self.emit(
            "image",
            item=item,
            metadata={
                "bytes": len(item.image_bytes or b""),
            },
        )

        self.sleep_simulated(
            self.rng.uniform(1.5, 4.0)
        )

        self.executor.enter(soft=False)
        self.sleep_simulated(
            self.rng.uniform(0.4, 1.0)
        )

    def paragraph_break(self, item: DocumentItem) -> None:
        self.executor.enter(soft=False)

        self.emit(
            "format",
            intended="\n\n",
            actual="\n\n",
            item=item,
            metadata={
                "action": "paragraph_break"
            },
        )

        self.sleep_simulated(
            self.rng.uniform(0.35, 0.9)
        )

        self.lines_since_break += 1

    def maybe_micro_break(self) -> None:
        if self.lines_since_break < self.next_break_target:
            return

        duration = self.rng.uniform(
            *SHORT_BREAK_MINUTES
        )

        self.emit(
            "break",
            metadata={
                "break_type": "short",
                "duration_minutes": duration,
            },
        )

        self.sleep_simulated(duration * 60)

        self.behavior.recover(duration)

        self.lines_since_break = 0
        self.next_break_target = self.rng.randint(
            *SHORT_BREAK_INTERVAL_LINES
        )

    # --------------------------------------------------------
    # SESSION
    # --------------------------------------------------------

    def run(self, start_index: int = 0) -> dict[str, Any]:
        session_started = datetime.now(timezone.utc)

        sprint_duration = self.rng.uniform(
            *WORK_SPRINT_MINUTES
        )
        sprint_elapsed = 0.0

        for index in range(start_index, len(self.items)):
            if self.elapsed_ms / 3_600_000 >= MAX_SESSION_HOURS:
                self.emit(
                    "session_stop",
                    metadata={
                        "reason": "max_session_hours"
                    },
                )
                break

            item = self.items[index]

            # Long break.
            if sprint_elapsed >= sprint_duration * 60:
                duration = self.rng.uniform(
                    *LONG_BREAK_MINUTES
                )

                self.emit(
                    "break",
                    item=item,
                    metadata={
                        "break_type": "long",
                        "duration_minutes": duration,
                    },
                )

                self.sleep_simulated(duration * 60)
                self.behavior.recover(duration)

                sprint_duration = self.rng.uniform(
                    *WORK_SPRINT_MINUTES
                )
                sprint_elapsed = 0.0

            self.apply_formatting(item)

            if item.item_type == "text":
                self.type_text(item)

            elif item.item_type == "soft_break":
                self.executor.enter(soft=True)
                self.emit(
                    "format",
                    intended="\n",
                    actual="\n",
                    item=item,
                    metadata={"action": "soft_break"},
                )
                self.sleep_simulated(
                    self.rng.uniform(0.25, 0.8)
                )

            elif item.item_type == "paragraph_break":
                self.paragraph_break(item)

            elif item.item_type == "image":
                self.insert_image(item)

            sprint_elapsed = self.elapsed_ms / 1000

            if (index + 1) % 25 == 0:
                self.state.save(index + 1, self.session_id, self.seed)

            self.maybe_micro_break()

        self.state.save(len(self.items), self.session_id, self.seed)

        # Turn formatting off at the end if needed.
        if self.current_formatting["bold"]:
            self.executor.toggle_bold()
            self.current_formatting["bold"] = False



        # Turn formatting off at the end if needed.
        if self.current_formatting["bold"]:
            self.executor.toggle_bold()
            self.current_formatting["bold"] = False

        summary = self.metrics.summary()

        summary.update({
            "session_id": self.session_id,
            "seed": self.seed,
            "profile": self.profile.name,
            "started_at": session_started.isoformat(),
            "finished_at": datetime.now(timezone.utc).isoformat(),
            "pdf": str(PDF_PATH),
        })

        self.config.summary_file.write_text(
            json.dumps(summary, indent=2),
            encoding="utf-8",
        )

        return summary


# ============================================================
# REPORTING
# ============================================================

def print_summary(summary: dict[str, Any]) -> None:
    print("\n" + "=" * 60)
    print("SIMULATION COMPLETE")
    print("=" * 60)

    fields = [
        ("Session", "session_id"),
        ("Seed", "seed"),
        ("Profile", "profile"),
        ("Duration (min)", "duration_minutes"),
        ("Characters", "characters"),
        ("Effective WPM", "effective_wpm"),
        ("Errors", "error_count"),
        ("Error rate", "error_rate"),
        ("Corrections", "correction_count"),
        ("Pauses", "pause_count"),
        ("Mean key latency (ms)", "mean_key_latency_ms"),
        ("Median key latency (ms)", "median_key_latency_ms"),
    ]

    for label, key in fields:
        print(f"{label:<28}: {summary.get(key)}")

    flags = summary.get("quality_flags", [])

    print("\nQuality flags:")
    if flags:
        for flag in flags:
            print(f"  - {flag}")
    else:
        print("  None")

    print("=" * 60)


# ============================================================
# MAIN
# ============================================================

def build_cli() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Empirically calibrated AI-training typing simulator"
    )
    parser.add_argument(
        "--reference-data",
        type=Path,
        default=REFERENCE_DATA_PATH,
        help="Local CSV/TSV reference keystroke dataset used for calibration.",
    )
    parser.add_argument(
        "--no-calibration",
        action="store_true",
        help="Disable empirical calibration even when reference data is configured.",
    )
    parser.add_argument(
        "--profile",
        choices=sorted(PROFILES),
        default=PROFILE_NAME,
        help="Human typing profile before empirical calibration.",
    )
    parser.add_argument(
        "--seed",
        type=int,
        default=RANDOM_SEED,
        help="Random seed. Use a fixed seed for reproducibility.",
    )
    parser.add_argument(
        "--offline",
        action="store_true",
        help="Generate the event stream without controlling the keyboard.",
    )
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

    if config.seed is None:
        seed = random.SystemRandom().randrange(0, 2**63 - 1)
    else:
        seed = config.seed

    state_manager = StateManager(config.state_file)
    state = state_manager.load()

    # A saved session keeps the same seed/session ID so it can be
    # reproduced after interruption.
    if state.get("seed") is not None:
        seed = int(state["seed"])

    session_id = state.get("session_id") or str(uuid.uuid4())
    start_index = int(state.get("item_index", 0))

    profile = PROFILES[config.profile_name]
    calibration_report = None

    # --------------------------------------------------------
    # EMPIRICAL CALIBRATION
    # --------------------------------------------------------
    reference_path = args.reference_data

    if ENABLE_EMPIRICAL_CALIBRATION and not args.no_calibration:
        if reference_path:
            print("Loading empirical reference data...")
            calibrator = EmpiricalCalibrator(reference_path)
            calibration_report = calibrator.fit()
            calibrator.write_report(
                calibration_report,
                CALIBRATION_REPORT_FILE,
            )
            profile = calibrator.apply_to_profile(
                profile,
                calibration_report,
            )
            print(
                "Empirical calibration applied: "
                f"{calibration_report.get('dataset_type')}"
            )
        else:
            print(
                "No empirical reference dataset configured. "
                "Using profile priors only."
            )

    print("=" * 60)
    print("AI TRAINING TYPING SIMULATOR")
    print("=" * 60)
    print(f"PDF:          {PDF_PATH}")
    print(f"Profile:      {config.profile_name}")
    print(f"Calibrated:   {calibration_report is not None}")
    print(f"Seed:         {seed}")
    print(f"Session:      {session_id}")
    print(f"Execution:    {config.execute}")
    print(f"Resume at:    item {start_index}")

    if calibration_report:
        iki = (
            calibration_report.get("inter_key_dd_seconds")
            or calibration_report.get("inter_key_seconds")
            or {}
        )
        print(
            "Reference IKI median: "
            f"{iki.get('median')} seconds"
        )
        print(
            "Calibrated base WPM:   "
            f"{profile.base_wpm:.2f}"
        )
        print(
            "Calibrated timing sigma: "
            f"{profile.key_interval_sigma:.3f}"
        )

    parser = PDFDocumentParser(PDF_PATH)
    items = parser.parse()

    if not items:
        raise RuntimeError("No readable content was found in the PDF.")

    print(f"Document items: {len(items)}")

    if config.execute:
        print(
            f"\nFocus the target application. "
            f"Simulation starts in {STARTUP_DELAY_SECONDS} seconds."
        )

        for remaining in range(
            STARTUP_DELAY_SECONDS,
            0,
            -1,
        ):
            print(
                f"Starting in {remaining}s...",
                end="\r",
                flush=True,
            )
            time.sleep(1)

        if human_click is not None:
            human_click(
                SAFE_TITLE_BAR_CLICK[0],
                SAFE_TITLE_BAR_CLICK[1],
            )

    simulation = TypingSimulation(
        items=items,
        config=config,
        profile=profile,
        seed=seed,
        session_id=session_id,
    )

    try:
        summary = simulation.run(
            start_index=start_index,
        )

        # Attach calibration provenance to the session summary without
        # copying the entire reference dataset into every event.
        if calibration_report:
            summary["calibration"] = {
                "enabled": True,
                "reference_file": str(reference_path),
                "dataset_type": calibration_report.get("dataset_type"),
                "rows_used": calibration_report.get("rows_used"),
                "calibration_report": str(CALIBRATION_REPORT_FILE),
            }
            config.summary_file.write_text(
                json.dumps(summary, indent=2),
                encoding="utf-8",
            )

    except KeyboardInterrupt:
        print("\nSimulation interrupted. State has been preserved.")
        raise

    print_summary(summary)

    # Successful completion removes resume state.
    if state_manager.path.exists():
        state_manager.clear()

    print(f"\nEvent log:        {EVENT_LOG_FILE}")
    print(f"Session summary:  {SESSION_SUMMARY_FILE}")
    if calibration_report:
        print(f"Calibration:      {CALIBRATION_REPORT_FILE}")


if __name__ == "__main__":
    main()
