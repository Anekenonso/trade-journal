# Trade Journal Maker

A Flask web app that reads MT4/MT5 trade history screenshots, extracts structured
trade data via OCR, and returns downloadable Excel journal files — without saving
any files on the server.

Live demo: https://trade-journal-dk31.onrender.com
*(Free-tier hosting: the app sleeps after inactivity, so the first request after
a while can take 30-60 seconds to wake up.)*

## What it does

- Upload one or more MT4/MT5 trade history screenshots
- Run OCR on each image to extract text
- Parse each trade block into structured fields: symbol, direction, lot size,
  entry price, exit price, stop loss, take profit, outcome, session, open/close
  date and time, and profit
- Automatically determine how each trade closed — **Take Profit**, **Stop Loss**,
  **Break Even**, or **Manual/Other** — by comparing the exit price against the
  stop loss/take profit levels and cross-checking against the realized profit
- Automatically tag each trade's **trading session** (Asian/Sydney, London,
  London/New York Overlap, New York) based on the open time shown on the screenshot
- Generate two in-memory Excel workbooks:
  - `trade_journal.xlsx` — one row per trade with all extracted fields
  - `trade_summary.xlsx` — win rate, net profit, outcome breakdown (TP/SL/BE/manual
    counts), session breakdown, and a chart
- Return both files together as `trade_journal_bundle.zip`

Non-trade screenshots (charts, other UI screens) are silently skipped rather than
producing garbage rows — if nothing in an upload matches a trade block, it
contributes zero rows instead of an error.

## How the parser works

The heuristic parser is built specifically around MT4/MT5's trade history layout,
where each closed position appears as a block like:

```
EURUSD, buy 0.25                    2026.08.10 11:13:43
1.15592 → 1.15508                   -21.00
#33890516            Open:          2026.08.10 07:54:38
S/L:      1.15511    Swap:          0.00
T/P:      1.15837    Commission:    -1.00
```

A few things worth knowing:

- **OCR mode matters.** Tesseract's default page segmentation reads this layout's
  left and right columns as two separate blocks (all left-column text top-to-bottom,
  then all right-column text), scrambling row order. The app forces `--psm 6`
  (uniform block of text) to read it in the correct row order.
- **Outcome classification** compares the exit price to the entry, stop loss, and
  take profit levels to determine which one the trade actually closed near, then
  sanity-checks that against the realized profit sign (a genuine stop-loss hit
  shouldn't show a profit, and vice versa) — this catches cases like a trailing
  stop or manual close that happens to land near the original SL/TP level.
- **Session tagging** is based on the *hour shown on the screenshot*, which is
  your broker/platform's server time — not necessarily GMT. If your broker uses
  a different UTC offset than expected, session labels will be off by that same
  offset.
- If OCR misses a field (blurry screenshot, cropped edge, unusual color theme),
  that field is left blank in the output rather than guessed.

## Optional AI-assisted parsing

If `OPENAI_API_KEY` is configured, the app will first try parsing the OCR text
with an OpenAI-compatible chat completion call before falling back to the
heuristic parser. This can produce cleaner results on messy OCR text.

**Status:** AgentRouter (a free/community OpenAI-compatible gateway, used here as
a no-cost alternative to a funded OpenAI account) currently sits behind an Aliyun
WAF that challenges server-to-server requests from hosts like Render with a
bot-verification page instead of returning API JSON. Because of this, AI parsing
is **not currently functional in the live deployment** — the app automatically
and silently falls back to the heuristic parser, which is fully functional on
its own and is what powers the live demo today.

To use AI-assisted parsing:

- **With OpenAI directly:** set `OPENAI_API_KEY` to a funded OpenAI key and leave
  `OPENAI_BASE_URL` unset.
- **With an OpenAI-compatible gateway:** set `OPENAI_API_KEY` to that provider's
  key and `OPENAI_BASE_URL` to their API base (e.g.
  `https://agentrouter.org/v1`). Confirm the gateway allows programmatic/server
  traffic before relying on it — some free gateways block datacenter IPs.

## Supported output fields

| Field | Description |
|---|---|
| Symbol | Traded instrument, e.g. `EURUSD` |
| Direction | `Buy` or `Sell` |
| Lot Size | Position size |
| Entry / Exit Price | Open and close price |
| Stop Loss / Take Profit | SL and TP levels set on the trade |
| Outcome | `Take Profit`, `Stop Loss`, `Break Even`, or `Manual/Other` |
| Session | `Asian/Sydney`, `London`, `London/New York Overlap`, or `New York` |
| Open/Close Date & Time | As shown on the screenshot |
| Profit | Realized profit/loss for the trade |

## App flow

- `GET /` — renders the upload page.
- `POST /parse` — returns a JSON preview of extracted trades and summary metrics.
- `POST /upload` and `POST /download` — return a zip bundle containing both
  Excel files.

## Setup

1. Install Python 3.11+.
2. Install dependencies:

```bash
python -m pip install -r requirements.txt
```

3. Install Tesseract OCR.
   - Windows: download and install from https://github.com/tesseract-ocr/tesseract
   - macOS: `brew install tesseract`
   - Linux: use your package manager, for example `sudo apt install tesseract-ocr`

4. Copy `.env.example` to `.env` if you want optional AI-assisted parsing:

```bash
cp .env.example .env
```

Then set:

```text
OPENAI_API_KEY=your_api_key_here
FLASK_SECRET_KEY=your-secret-key

# Optional: only needed if using a gateway other than OpenAI directly
# OPENAI_BASE_URL=https://agentrouter.org/v1
# OPENAI_MODEL=gpt-4o-mini
```

## Run locally

```bash
python app.py
```

Open `http://127.0.0.1:5000` in your browser.

## Run with Docker

The included `Dockerfile` installs Tesseract and runs the app with gunicorn —
this is how the live deployment on Render runs.

```bash
docker build -t trade-journal-maker .
docker run -p 7860:7860 --env-file .env trade-journal-maker
```

Note: the container defaults to `--workers 1 --threads 2` for gunicorn. This is
intentional — free-tier hosts like Render's free plan have limited RAM (512MB),
and running multiple worker processes was found to cause out-of-memory kills
under this stack's footprint (Flask + pandas + openpyxl + Pillow + a Tesseract
subprocess). If deploying with more available memory, workers can be increased.

## Usage

- Click `Preview Trades` after uploading screenshots to verify extraction results
  before downloading.
- Click `Download ZIP` to get both Excel files.
- No files are stored on the server: everything is generated and streamed in memory.
- Uploads can mix trade-history screenshots with unrelated images (like chart
  views) — only screenshots matching the MT4/MT5 trade block pattern contribute
  rows.

## Notes

- The preview uses an AJAX request, so the page does not need to refresh.
- The download endpoint packages the files into `trade_journal_bundle.zip`.
- Large screenshots are automatically downscaled (max 2000px on the longest
  side) before OCR to keep memory usage reasonable on low-RAM hosts.
- If OCR misses values, try a higher-resolution or higher-contrast screenshot.

## Troubleshooting

- **"No valid trade data could be extracted"** — verify the screenshot is an
  MT4/MT5 trade history view (not a chart or other screen), and that it clearly
  shows the `SYMBOL, buy/sell lots`, `S/L:`, and `T/P:` labels.
- **Tesseract not found** — make sure the Tesseract executable is installed and
  on your PATH (local dev) or that the Docker image built successfully (hosted).
- **AI parsing always falls back to heuristics** — check the application logs
  for the specific warning. A `response_type=str` with HTML/WAF content means
  the configured gateway is blocking the request before it reaches the actual
  API — see "Optional AI-assisted parsing" above. Other errors (auth, quota,
  model-not-found) will show a clearer message from the API itself.
- **Out-of-memory / worker SIGKILL on free hosting tiers** — reduce gunicorn
  workers to 1 (already the default in the provided Dockerfile) and keep
  uploaded images to reasonable sizes.
