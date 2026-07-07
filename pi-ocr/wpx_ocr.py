#!/usr/bin/env python3
"""
WPX screenshot -> CSV OCR pipeline.

Flow:
  1. Pull new screenshots from Dropbox (dbx:"WPX Screenshots")
  2. OCR each with Tesseract (after light image cleanup)
  3. Parse "player name  score" lines into rows
  4. Write one CSV per screenshot  ->  <name>.csv
  5. Push CSVs back to Dropbox (dbx:"Score CSVs")
  6. Remember processed files so they aren't re-done

Run:  source ~/ocr/venv/bin/activate && python3 ~/ocr/wpx_ocr.py
"""

import csv
import re
import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageOps

# ── CONFIG ────────────────────────────────────────────────────────────
REMOTE          = "dbx"
INBOX_REMOTE    = f'{REMOTE}:"WPX Screenshots"'
OUTPUT_REMOTE   = f'{REMOTE}:"Score CSVs"'

BASE      = Path.home() / "ocr"
INBOX     = BASE / "inbox"        # downloaded screenshots
CSV_OUT   = BASE / "csv"          # generated CSVs
DONE_LOG  = BASE / "processed.txt"  # filenames already handled

IMG_EXTS  = {".png", ".jpg", ".jpeg", ".webp"}

# A score is the trailing number on a row (0, or grouped like 19,831,527).
SCORE_RE   = re.compile(r"(\d[\d.,\s]*)\s*$")
# Alliance tag / subtitle text to strip out of the player name.
TAG_RE     = re.compile(r"\[?wpx\]?\s*warrior|phoenix", re.IGNORECASE)
RANK_RE    = re.compile(r"^\d{1,3}[\).:]?\s+")


# ── HELPERS ───────────────────────────────────────────────────────────
def run(cmd):
    """Run a shell command, return stdout, raise on failure."""
    res = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    if res.returncode != 0:
        print(f"  ! command failed: {cmd}\n    {res.stderr.strip()}", file=sys.stderr)
    return res.stdout


def load_done():
    if DONE_LOG.exists():
        return set(DONE_LOG.read_text().splitlines())
    return set()


def mark_done(name):
    with DONE_LOG.open("a") as f:
        f.write(name + "\n")


THRESHOLD = 140   # binarization cutoff; lower keeps more (darker) pixels


def preprocess(img_path):
    """Grayscale, 2x upscale, binarize -> best OCR on game scoreboards."""
    img = Image.open(img_path).convert("L")
    w, h = img.size
    img = img.resize((w * 2, h * 2), Image.LANCZOS)   # upscale helps small text
    img = img.point(lambda p: 255 if p > THRESHOLD else 0)   # black/white
    out = img_path.with_suffix(".prep.png")
    img.save(out)
    return out


def ocr(img_path):
    """Return raw OCR text for an image."""
    prepped = preprocess(img_path)
    # --psm 6: treat the image as a single uniform block of text (one row per line)
    txt = run(f'tesseract "{prepped}" stdout --psm 6')
    prepped.unlink(missing_ok=True)
    return txt


def parse_scores(text):
    """Turn raw OCR text into [(name, score:int)] rows."""
    rows = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        m = SCORE_RE.search(line)
        if not m:
            continue
        raw_score = re.sub(r"[.,\s]", "", m.group(1))
        if not raw_score.isdigit():
            continue
        score = int(raw_score)
        name = line[: m.start()]
        name = RANK_RE.sub("", name)          # drop leading rank "94 "
        name = TAG_RE.sub("", name)           # drop "[Wpx]Warrior Phoenix"
        name = name.strip(" .-:|\t")
        if name:
            rows.append((name, score))
    return rows


def write_csv(rows, dest):
    with dest.open("w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Player", "Score"])
        w.writerows(rows)


# ── MAIN ──────────────────────────────────────────────────────────────
def main():
    for d in (INBOX, CSV_OUT):
        d.mkdir(parents=True, exist_ok=True)

    print("Pulling screenshots from Dropbox...")
    run(f'rclone copy {INBOX_REMOTE} "{INBOX}" --progress')

    done = load_done()
    images = [p for p in sorted(INBOX.iterdir()) if p.suffix.lower() in IMG_EXTS]
    new = [p for p in images if p.name not in done]

    if not new:
        print("No new screenshots to process.")
        return

    for img in new:
        print(f"\nProcessing {img.name} ...")
        text = ocr(img)
        rows = parse_scores(text)
        print(f"  found {len(rows)} score rows")
        csv_path = CSV_OUT / (img.stem + ".csv")
        write_csv(rows, csv_path)
        # also drop the raw OCR text next to it for spot-checking
        (CSV_OUT / (img.stem + ".txt")).write_text(text)
        mark_done(img.name)

    print("\nPushing CSVs back to Dropbox...")
    run(f'rclone copy "{CSV_OUT}" {OUTPUT_REMOTE} --include "*.csv" --progress')
    print("Done.")


if __name__ == "__main__":
    main()
