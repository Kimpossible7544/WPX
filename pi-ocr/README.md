# pi-ocr — free screenshot → CSV pipeline

A Raspberry Pi replacement for the (cost-disabled) AI screenshot processing in
`WPX-scores.html`. Instead of sending scoreboard screenshots to an AI model, it
runs [Tesseract](https://github.com/tesseract-ocr/tesseract) locally on the Pi.

## Flow

```
Dropbox: "WPX Screenshots"   (drop screenshots here)
        └── rclone copy → Pi
                └── preprocess (grayscale, 2x upscale, binarize)
                        └── Tesseract OCR  (--psm 6)
                                └── parse "player  score" rows
                                        └── one CSV per screenshot (Player,Score)
                                                └── rclone copy → Dropbox: "Score CSVs"
```

Processed filenames are logged to `~/ocr/processed.txt` so re-runs skip them.

## Setup on the Pi

```bash
sudo apt update
sudo apt install -y tesseract-ocr python3-venv rclone
mkdir -p ~/ocr && cd ~/ocr
python3 -m venv venv && source venv/bin/activate
pip install pillow
# configure a Dropbox remote named "dbx"
rclone config          # headless: choose "no" for auto config, authorize on a PC
```

Copy `wpx_ocr.py` into `~/ocr/` and run:

```bash
source ~/ocr/venv/bin/activate
python3 ~/ocr/wpx_ocr.py
```

To automate every hour, add to `crontab -e`:

```
0 * * * * /home/USER/ocr/venv/bin/python /home/USER/ocr/wpx_ocr.py >> /home/USER/ocr/run.log 2>&1
```

## Accuracy caveats

Tesseract is far less accurate than the AI model on stylized game UI. Expect to
verify output. The raw OCR text for each screenshot is saved next to its CSV
(`<name>.txt`) for spot-checking. Tuning knobs live at the top of `wpx_ocr.py`:

- `THRESHOLD` — binarization cutoff (lower keeps more dark pixels).
- Tesseract `--psm` mode in `ocr()` (`6` = uniform block; try `4` or `11`).
- `SCORE_RE` / `TAG_RE` — row parsing regexes.

Cropping screenshots to just the name+score region before OCR improves results.
