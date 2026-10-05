#!/usr/bin/env python3
"""
boost_fps.py - Raise the frame rate of every video in a folder to at least a target fps (default 30).

Videos already at or above the target are skipped (or copied with --copy-ok).
Originals are never modified; results go to an output folder.

Requires FFmpeg and ffprobe on your PATH.

Examples:
    python boost_fps.py "D:\\videos"
    python boost_fps.py ./videos --method interpolate --crf 20
    python boost_fps.py ./videos -o ./converted --recursive --copy-ok
"""

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

VIDEO_EXTS = {".mp4", ".mov", ".mkv", ".avi", ".webm", ".m4v", ".flv", ".wmv", ".mpg", ".mpeg", ".ts"}


def parse_rate(rate: str) -> float:
    """Turn an ffprobe rate string like '30000/1001' into a float."""
    try:
        if "/" in rate:
            num, den = rate.split("/")
            den = float(den)
            return float(num) / den if den else 0.0
        return float(rate)
    except (ValueError, ZeroDivisionError):
        return 0.0


def get_fps(path: Path) -> float:
    """Return the average fps of the first video stream (0.0 if unreadable)."""
    result = subprocess.run(
        [
            "ffprobe", "-v", "error",
            "-select_streams", "v:0",
            "-show_entries", "stream=avg_frame_rate,r_frame_rate",
            "-of", "default=noprint_wrappers=1:nokey=1",
            str(path),
        ],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        return 0.0
    rates = [parse_rate(line.strip()) for line in result.stdout.splitlines() if line.strip()]
    # avg_frame_rate comes first; fall back to r_frame_rate if it's 0
    for r in rates:
        if r > 0:
            return r
    return 0.0


def build_filter(method: str, target: float) -> str:
    if method == "interpolate":
        return f"minterpolate=fps={target}:mi_mode=mci:mc_mode=aobmc:vsbmc=1"
    return f"fps={target}"


def run_ffmpeg(src: Path, dst: Path, vf: str, crf: int, preset: str, audio_codec: str) -> bool:
    cmd = [
        "ffmpeg", "-hide_banner", "-loglevel", "error", "-y",
        "-i", str(src),
        "-map", "0:v:0", "-map", "0:a?",
        "-vf", vf,
        "-c:v", "libx264", "-crf", str(crf), "-preset", preset,
        "-pix_fmt", "yuv420p",
        "-c:a", audio_codec,
        "-movflags", "+faststart",
        str(dst),
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        err = result.stderr.strip().splitlines()
        if err:
            print(f"      ffmpeg: {err[-1]}")
        return False
    return True


def convert(src: Path, dst: Path, vf: str, crf: int, preset: str) -> bool:
    dst.parent.mkdir(parents=True, exist_ok=True)
    # Try to keep the original audio untouched; fall back to AAC if the container can't hold it.
    if run_ffmpeg(src, dst, vf, crf, preset, "copy"):
        return True
    if dst.exists():
        dst.unlink()
    print("      retrying with AAC audio...")
    ok = run_ffmpeg(src, dst, vf, crf, preset, "aac")
    if not ok and dst.exists():
        dst.unlink()
    return ok


def find_videos(folder: Path, recursive: bool, exclude: Path):
    pattern = "**/*" if recursive else "*"
    for p in sorted(folder.glob(pattern)):
        if not p.is_file() or p.suffix.lower() not in VIDEO_EXTS:
            continue
        try:
            p.relative_to(exclude)
            continue  # inside the output folder - skip
        except ValueError:
            yield p


def main() -> int:
    ap = argparse.ArgumentParser(description="Raise video frame rate to at least a target fps.")
    ap.add_argument("folder", type=Path, help="Folder containing the videos")
    ap.add_argument("-o", "--output", type=Path, default=None,
                    help="Output folder (default: <folder>/fps_boosted)")
    ap.add_argument("--fps", type=float, default=30.0, help="Target fps (default: 30)")
    ap.add_argument("--method", choices=["duplicate", "interpolate"], default="duplicate",
                    help="duplicate = fast, repeats frames; interpolate = slow, smoother (default: duplicate)")
    ap.add_argument("--crf", type=int, default=18, help="x264 quality, lower = better/larger (default: 18)")
    ap.add_argument("--preset", default="medium",
                    help="x264 speed preset: ultrafast..veryslow (default: medium)")
    ap.add_argument("--tolerance", type=float, default=0.0,
                    help="Treat videos within this many fps below the target as OK, "
                         "e.g. 0.1 skips 29.97fps files (default: 0)")
    ap.add_argument("-r", "--recursive", action="store_true", help="Include subfolders")
    ap.add_argument("--copy-ok", action="store_true",
                    help="Copy videos that already meet the target into the output folder")
    ap.add_argument("--overwrite", action="store_true", help="Redo files that already exist in the output folder")
    args = ap.parse_args()

    if not shutil.which("ffmpeg") or not shutil.which("ffprobe"):
        print("Error: ffmpeg and ffprobe must be installed and on your PATH.")
        return 1

    folder = args.folder.resolve()
    if not folder.is_dir():
        print(f"Error: '{folder}' is not a folder.")
        return 1

    out_dir = (args.output or folder / "fps_boosted").resolve()
    vf = build_filter(args.method, args.fps)
    threshold = args.fps - args.tolerance

    videos = list(find_videos(folder, args.recursive, out_dir))
    if not videos:
        print("No video files found.")
        return 0

    print(f"Found {len(videos)} video(s). Target: {args.fps:g} fps | method: {args.method} | output: {out_dir}\n")

    converted = skipped = failed = 0
    for i, src in enumerate(videos, 1):
        rel = src.relative_to(folder)
        dst = (out_dir / rel).with_suffix(".mp4")
        tag = f"[{i}/{len(videos)}] {rel}"

        fps = get_fps(src)
        if fps <= 0:
            print(f"{tag}: could not read fps - skipped")
            failed += 1
            continue

        if fps >= threshold:
            if args.copy_ok:
                dst_copy = out_dir / rel
                dst_copy.parent.mkdir(parents=True, exist_ok=True)
                if not dst_copy.exists() or args.overwrite:
                    shutil.copy2(src, dst_copy)
                print(f"{tag}: {fps:.2f} fps - already OK, copied")
            else:
                print(f"{tag}: {fps:.2f} fps - already OK, skipped")
            skipped += 1
            continue

        if dst.exists() and not args.overwrite:
            print(f"{tag}: output exists - skipped (use --overwrite to redo)")
            skipped += 1
            continue

        print(f"{tag}: {fps:.2f} fps -> {args.fps:g} fps ...")
        if convert(src, dst, vf, args.crf, args.preset):
            converted += 1
        else:
            print("      FAILED")
            failed += 1

    print(f"\nDone. Converted: {converted} | Skipped: {skipped} | Failed: {failed}")
    return 0 if failed == 0 else 2


if __name__ == "__main__":
    sys.exit(main())