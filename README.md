# WPX Alliance Score Tracker

A full-stack web application for tracking and analyzing player performance data for a competitive mobile gaming alliance. Built entirely from scratch using vanilla JavaScript, HTML, and CSS — no frameworks, no build tools.

🔗 **[Live Site](https://kimpossible7544.github.io/WPX/)**

---

## What It Does

Tracks daily and weekly scores for a 96-member alliance, providing real-time standings, performance trends, and goal tracking. The system processes raw score data, matches player identities across name variants, and presents the results in a clean dashboard interface.

### Features

- **Daily score ingestion** — processes score data and maps in-game display names to canonical roster names using fuzzy matching and alias lookup
- **Weekly goal tracking** — flags missed daily goals (sub-6M) and missed weekly goals (sub-20M) per player
- **Tyrant event tracking** — Y/N/F scoring with threshold-based pass/fail logic
- **Active vs. Resting week logic** — performance metrics automatically exclude resting weeks from goal miss counts and rankings
- **Live data pipeline** — pulls data from a Dropbox-hosted Excel workbook via CDN using SheetJS, no server required
- **Responsive dashboard** — works on desktop and mobile

---

## Tech Stack

| Layer | Technology |
|---|---|
| Frontend | HTML, CSS, JavaScript (vanilla) |
| Data parsing | SheetJS |
| Data hosting | Dropbox CDN |
| Deployment | GitHub Pages |
| Authentication | Custom roster-ID based login system |

---

## Architecture

```
Excel Workbook (.xlsm)
    └── Dropbox (public CDN link)
            └── SheetJS (client-side parser)
                    └── wpx_data.js (data pipeline)
                            └── GitHub Pages dashboard
```

The Excel workbook serves as the backend database. A VBA macro system handles score imports, player ID normalization, and weekly calculations. The frontend fetches the workbook directly from Dropbox on page load and parses it entirely in the browser — no API, no server, no build step.

---

## Key Technical Challenges Solved

- **Name normalization** — players use special characters, Unicode variants, and alternate display names. The matching system strips, normalizes, and maps names across Cyrillic, Greek, and Latin character sets
- **Data integrity** — duplicate ID detection, ambiguous match flagging, and graceful fallback handling
- **Resting week exclusions** — week-level flags propagate through all derived metrics (averages, miss counts, rankings) without affecting raw score totals

---

## Related

The score tracker frontend is backed by a sophisticated Excel/VBA workbook system — see the [Player-Dashboard](https://github.com/Kimpossible7544/Player-Dashboard) repo for the analytics layer.
