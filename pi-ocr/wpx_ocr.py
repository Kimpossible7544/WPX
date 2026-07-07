#!/usr/bin/env python3
"""
WPX screenshot -> CSV OCR pipeline (EasyOCR).

Flow:
  1. Pull new screenshots from Dropbox (dbx:"WPX Screenshots")
  2. Run EasyOCR on each -> text boxes with x/y positions
  3. Reconstruct scoreboard rows spatially: rank (left col),
     player name (middle col, minus the "[Wpx]Warrior Phoenix" tag),
     score (right col)
  4. Write one CSV per screenshot -> Rank,Player,Score
  5. Push CSVs back to Dropbox (dbx:"Score CSVs")
  6. Remember processed files so they aren't re-done

Run:  source ~/ocr/venv/bin/activate && python3 ~/ocr/wpx_ocr.py

EasyOCR is far more accurate than Tesseract on stylized game UI, but the score
column can still be misread by a digit or two, and light "0" scores may not be
detected at all (left blank). The raw detections are saved next to each CSV
(<name>.txt) for spot-checking.
"""

import csv
import re
import subprocess
import sys
from pathlib import Path

import easyocr
from PIL import Image

# ── CONFIG ────────────────────────────────────────────────────────────
REMOTE          = "dbx"
INBOX_REMOTE    = f'{REMOTE}:"WPX Screenshots"'
OUTPUT_REMOTE   = f'{REMOTE}:"Score CSVs"'

BASE      = Path.home() / "ocr"
INBOX     = BASE / "inbox"          # downloaded screenshots
CSV_OUT   = BASE / "csv"            # generated CSVs
DONE_LOG  = BASE / "processed.txt"  # filenames already handled

IMG_EXTS  = {".png", ".jpg", ".jpeg", ".webp"}

# Column boundaries as a fraction of image width (works across resolutions).
RANK_COL_MAX   = 0.25   # rank number sits left of this
SCORE_COL_MIN  = 0.60   # score sits right of this
# A name/score belongs to a rank if within this fraction of image height in y.
ROW_BAND       = 0.055

# Alliance tag / subtitle text to drop from the player name.
TAG_RE = re.compile(r"\[?wpx\]?\s*warrior|phoenix", re.IGNORECASE)
# UI chrome (tabs, day selector, buttons) that must never be read as a name.
HEADER_RE = re.compile(
    r"^(mon|tue|wed|thu|fri|sat|sun|ranking|this week'?s|daily rank|"
    r"personal ranking|my alliance)$",
    re.IGNORECASE,
)

# Load the OCR model once (English). gpu=False for the Pi.
_reader = None


def reader():
    global _reader
    if _reader is None:
        print("Loading EasyOCR model (first run downloads it)...")
        _reader = easyocr.Reader(["en"], gpu=False)
    return _reader


# ── HELPERS ───────────────────────────────────────────────────────────
def run(cmd):
    res = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    if res.returncode != 0:
        print(f"  ! command failed: {cmd}\n    {res.stderr.strip()}", file=sys.stderr)
    return res.stdout


def load_done():
    return set(DONE_LOG.read_text().splitlines()) if DONE_LOG.exists() else set()


def mark_done(name):
    with DONE_LOG.open("a") as f:
        f.write(name + "\n")


def as_number(text):
    """Return int if text is a number (allowing , . space grouping), else None."""
    stripped = re.sub(r"[.,\s]", "", text)
    return int(stripped) if stripped.isdigit() else None


def detect(img_path):
    """Return (list of boxes, image width, image height)."""
    w, h = Image.open(img_path).size
    raw = reader().readtext(str(img_path), detail=1, paragraph=False)
    items = []
    for box, text, conf in raw:
        cx = sum(p[0] for p in box) / 4
        cy = sum(p[1] for p in box) / 4
        items.append({
            "cx": cx, "cy": cy, "text": text.strip(), "conf": conf,
            "num": as_number(text), "is_tag": bool(TAG_RE.search(text)),
            "is_header": bool(HEADER_RE.match(text.strip())),
        })
    return items, w, h


def build_rows(items, w, h):
    """Reconstruct [(rank, name, score)] from positioned OCR boxes."""
    rank_max  = RANK_COL_MAX * w
    score_min = SCORE_COL_MIN * w
    band      = ROW_BAND * h

    ranks = sorted(
        (it for it in items if it["num"] is not None and it["cx"] < rank_max),
        key=lambda it: it["cy"],
    )

    rows = []
    for r in ranks:
        y = r["cy"]

        name_cands = [
            it for it in items
            if rank_max <= it["cx"] <= score_min
            and not it["is_tag"] and not it["is_header"]
            and it["text"] and it["num"] is None
            and abs(it["cy"] - y) < band
        ]
        name = min(name_cands, key=lambda it: abs(it["cy"] - y))["text"] if name_cands else ""

        score_cands = [
            it for it in items
            if it["cx"] >= score_min and it["num"] is not None
            and abs(it["cy"] - y) < band
        ]
        score = min(score_cands, key=lambda it: abs(it["cy"] - y))["num"] if score_cands else None

        if name or score is not None:
            rows.append((r["num"], name, score))
    return rows


def write_csv(rows, dest):
    with dest.open("w", newline="") as f:
        wr = csv.writer(f)
        wr.writerow(["Rank", "Player", "Score"])
        for rank, name, score in rows:
            wr.writerow([rank, name, "" if score is None else score])


def dump_raw(items, dest):
    lines = [f"{it['conf']:.2f}  x={int(it['cx']):4d} y={int(it['cy']):4d}  {it['text']}"
             for it in sorted(items, key=lambda i: i["cy"])]
    dest.write_text("\n".join(lines))


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
        items, w, h = detect(img)
        rows = build_rows(items, w, h)
        print(f"  found {len(rows)} player rows")
        write_csv(rows, CSV_OUT / (img.stem + ".csv"))
        dump_raw(items, CSV_OUT / (img.stem + ".txt"))
        mark_done(img.name)

    print("\nPushing CSVs back to Dropbox...")
    run(f'rclone copy "{CSV_OUT}" {OUTPUT_REMOTE} --include "*.csv" --progress')
    print("Done.")


if __name__ == "__main__":
    main()
