# pi-ocr — free screenshot → CSV pipeline

A Raspberry Pi replacement for the (cost-disabled) AI screenshot processing in
`WPX-scores.html`. Instead of sending scoreboard screenshots to an AI model, it
runs [EasyOCR](https://github.com/JaidedAI/EasyOCR) locally on the Pi and
reconstructs each row spatially (rank / player / score) from the detected text
boxes.

## Flow

```
Dropbox: "WPX Screenshots"   (drop screenshots here)
        └── rclone copy → Pi
                └── EasyOCR → positioned text boxes
                        └── reconstruct rows by column (rank | name | score)
                                └── one CSV per screenshot (Rank,Player,Score)
                                        └── rclone copy → Dropbox: "Score CSVs"
```

Processed filenames are logged to `~/ocr/processed.txt` so re-runs skip them.

## Setup on the Pi

```bash
sudo apt update
sudo apt install -y python3-venv rclone libgl1 libglib2.0-0   # libGL etc. for opencv
mkdir -p ~/ocr && cd ~/ocr
python3 -m venv venv && source venv/bin/activate
pip install --upgrade pip
pip install easyocr pillow
# configure a Dropbox remote named "dbx"
rclone config          # headless: choose "no" for auto config, authorize on a PC
```

> **Note:** EasyOCR pulls in PyTorch, which is a large install and slow on a Pi.
> Use a Pi 4/5 with plenty of free storage. The **first run** also downloads the
> detection + recognition models (~65 MB) into `~/.EasyOCR`.

Copy `wpx_ocr.py` into `~/ocr/` and run:

```bash
source ~/ocr/venv/bin/activate
python3 ~/ocr/wpx_ocr.py
```

To automate every hour, add to `crontab -e`:

```
0 * * * * /home/USER/ocr/venv/bin/python /home/USER/ocr/wpx_ocr.py >> /home/USER/ocr/run.log 2>&1
```

## Accuracy

Tested on a real "Personal Ranking" screenshot, EasyOCR read player names and the
score column reliably (`DestinedOnè`, `BABY FORTUNER`, `Kimpossible7544`, etc.),
where plain Tesseract failed. Caveats:

- The score column can still be off by a digit on large numbers
  (e.g. `19,831,527` was read as `19,8315527`) — spot-check big values.
- Very light "0" scores are sometimes not detected and are left blank.
- Names with heavy styling/spacing come through approximately (`L i s a°`,
  `Ashmit Farm 1`) — the dashboard's fuzzy roster matching handles these.

The raw detections (text + x/y + confidence) for each screenshot are saved next
to its CSV as `<name>.txt` for verification.

## Tuning knobs (top of `wpx_ocr.py`)

- `RANK_COL_MAX` / `SCORE_COL_MIN` — column boundaries as a fraction of image
  width. Adjust if your screenshots put rank/score in different positions.
- `ROW_BAND` — how close (fraction of image height) a name/score must be to a
  rank to belong to the same row.
- `TAG_RE` — alliance-tag text stripped from names.
- `HEADER_RE` — UI chrome (tabs, day selector) ignored so it isn't read as a name.
