import io
import json
import os
import re
import zipfile
from statistics import mean

from flask import Flask, flash, jsonify, redirect, render_template, request, send_file, url_for
from PIL import Image
import pandas as pd
import pytesseract
from pytesseract import TesseractNotFoundError
from openpyxl.chart import BarChart, Reference

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    import openai
except ImportError:
    openai = None

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "change-me")

ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "webp", "bmp", "tiff"}
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")

openai_client = None
if OPENAI_API_KEY and openai:
    openai_client = openai.OpenAI(api_key=OPENAI_API_KEY)


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def ocr_image(image: Image.Image) -> str:
    # Downscale very large screenshots before OCR to keep memory usage
    # in check on low-RAM hosts. Tesseract doesn't need full resolution.
    max_dimension = 2000
    if max(image.size) > max_dimension:
        image.thumbnail((max_dimension, max_dimension), Image.LANCZOS)

    try:
        return pytesseract.image_to_string(image, lang="eng", config="--psm 6")
    except TesseractNotFoundError as exc:
        raise RuntimeError(
            "Tesseract OCR executable not found. Please install Tesseract and add it to your PATH. "
            "See README for installation instructions."
        ) from exc


SYMBOL_LINE_RE = re.compile(
    r'\b([A-Z]{2,10}(?:\.[A-Za-z]+)?)\s*,\s*(buy|sell)\s+([\d]*\.?[\d]+)',
    re.IGNORECASE,
)

PRICE_ARROW_PROFIT_RE = re.compile(
    r'([\d]+\.\d+)\s*(?:\u2192|\u2014|\u2013|->|-{1,2}>|>|~)\s*([\d]+\.\d+)[^\d+\-\n]{0,40}([+-]?\d+\.\d{1,2})?'
)

# Fallback for when OCR drops the arrow character entirely, leaving two
# bare decimal numbers (optionally followed by a signed profit figure).
PRICE_PAIR_FALLBACK_RE = re.compile(
    r'^\s*([\d]+\.\d+)\s+([\d]+\.\d+)\s*([+-]?\d+\.\d{1,2})?\s*$',
    re.MULTILINE,
)

SL_RE = re.compile(r'S\s*/\s*L\s*:?\s*([\d]+\.\d+)', re.IGNORECASE)
TP_RE = re.compile(r'T\s*/\s*P\s*:?\s*([\d]+\.\d+)', re.IGNORECASE)
DATETIME_RE = re.compile(r'(\d{4}[./]\d{2}[./]\d{2})\s+(\d{1,2}:\d{2}(?::\d{2})?)')

# Session windows keyed to the hour shown on the screenshot (broker/server
# time as displayed by the platform, not necessarily GMT).
SESSION_WINDOWS = [
    (0, 7, "Asian/Sydney"),
    (7, 12, "London"),
    (12, 16, "London/New York Overlap"),
    (16, 21, "New York"),
    (21, 24, "Asian/Sydney"),
]


def get_trading_session(time_str: str) -> str:
    if not time_str:
        return ""
    try:
        hour = int(time_str.split(":")[0])
    except (ValueError, IndexError):
        return ""
    for start, end, name in SESSION_WINDOWS:
        if start <= hour < end:
            return name
    return ""


def determine_outcome(entry, exit_price, sl, tp, profit=None) -> str:
    if entry is None or exit_price is None:
        return ""

    candidates = {"Break Even": abs(exit_price - entry)}
    if sl is not None:
        candidates["Stop Loss"] = abs(exit_price - sl)
    if tp is not None:
        candidates["Take Profit"] = abs(exit_price - tp)

    label = min(candidates, key=candidates.get)
    nearest_distance = candidates[label]

    if label != "Break Even":
        reference_price = sl if label == "Stop Loss" else tp
        reference_distance = abs(entry - reference_price)
        if reference_distance != 0 and nearest_distance > reference_distance * 0.5:
            label = "Manual/Other"

    # Sanity check against the realized profit: a genuine stop-loss hit
    # shouldn't show a profit, and a genuine take-profit hit shouldn't show
    # a loss. If it does, the price landed near that level for another
    # reason (trailing stop, manual close), so don't mislabel it.
    if profit is not None:
        if label == "Stop Loss" and profit > 0:
            label = "Manual/Other"
        elif label == "Take Profit" and profit < 0:
            label = "Manual/Other"

    return label


def split_trade_blocks(text: str) -> list[str]:
    matches = list(SYMBOL_LINE_RE.finditer(text))
    blocks = []
    for i, match in enumerate(matches):
        start = match.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        blocks.append(text[start:end])
    return blocks


def parse_ocr_text_simple(text: str) -> list[dict]:
    """Parse MT4/MT5 trade history screenshot text into structured rows.

    Targets blocks shaped like:
        EURUSD, buy 0.25                    2026.08.10 11:13:43
        1.15592 -> 1.15508                  -21.00
        #33890516            Open:          2026.08.10 07:54:38
        S/L:      1.15511    Swap:          0.00
        T/P:      1.15837    Commission:    -1.00
    """
    entries = []
    blocks = split_trade_blocks(text)

    for block in blocks:
        symbol_match = SYMBOL_LINE_RE.search(block)
        if not symbol_match:
            continue

        symbol = symbol_match.group(1).upper()
        direction = symbol_match.group(2).capitalize()
        lot_size = symbol_match.group(3)

        price_match = PRICE_ARROW_PROFIT_RE.search(block)
        if not price_match:
            price_match = PRICE_PAIR_FALLBACK_RE.search(block)

        entry_str = price_match.group(1) if price_match else ""
        exit_str = price_match.group(2) if price_match else ""
        profit = price_match.group(3) if price_match and price_match.group(3) else ""

        entry_price = parse_price(entry_str)
        exit_price = parse_price(exit_str)

        sl_match = SL_RE.search(block)
        tp_match = TP_RE.search(block)
        sl_price = parse_price(sl_match.group(1)) if sl_match else None
        tp_price = parse_price(tp_match.group(1)) if tp_match else None

        datetimes = DATETIME_RE.findall(block)
        close_date, close_time = datetimes[0] if len(datetimes) > 0 else ("", "")
        open_date, open_time = datetimes[1] if len(datetimes) > 1 else ("", "")

        outcome = determine_outcome(entry_price, exit_price, sl_price, tp_price, parse_price(profit))
        session = get_trading_session(open_time)

        entries.append(
            {
                "symbol": symbol,
                "direction": direction,
                "lot_size": lot_size,
                "entry": entry_str,
                "exit_price": exit_str,
                "stop_loss": sl_match.group(1) if sl_match else "",
                "take_profit": tp_match.group(1) if tp_match else "",
                "outcome": outcome,
                "session": session,
                "open_date": open_date,
                "open_time": open_time,
                "close_date": close_date,
                "close_time": close_time,
                "profit": profit,
            }
        )

    return entries


def parse_text_to_trades(text: str) -> list[dict]:
    if openai_client:
        prompt = (
            "Extract structured trading journal entries from this MT4/MT5 trade history "
            "screenshot text. Return a JSON array of objects with exactly these fields: "
            "symbol, direction (Buy or Sell), lot_size, entry, exit_price, stop_loss, "
            "take_profit, outcome, session, open_date, open_time, close_date, close_time, profit. "
            "For outcome, compare the exit price to stop_loss and take_profit: if exit is at or "
            "near take_profit, use 'Take Profit'; if at or near stop_loss, use 'Stop Loss'; if at "
            "or near the entry price (trader likely moved stop loss to break even), use "
            "'Break Even'; otherwise use 'Manual/Other'. "
            "For session, use the open_time hour against these broker-time windows: "
            "00:00-07:00 Asian/Sydney, 07:00-12:00 London, 12:00-16:00 London/New York Overlap, "
            "16:00-21:00 New York, 21:00-24:00 Asian/Sydney. "
            "If a field is missing from the text, return an empty string. Do not include any extra keys. "
            f"Text:\n{text.strip()}"
        )

        try:
            response = openai_client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.2,
                max_tokens=600,
            )
            content = response.choices[0].message.content.strip()
            cleaned = content
            if cleaned.startswith("```json") and cleaned.endswith("```"):
                cleaned = cleaned[7:-3].strip()
            return json.loads(cleaned)
        except Exception as exc:
            app.logger.warning("OpenAI parsing failed, falling back to heuristics: %s", exc)

    return parse_ocr_text_simple(text)


def parse_uploaded_files(files) -> list[dict]:
    trade_rows = []
    for uploaded_file in files:
        if uploaded_file.filename == "" or not allowed_file(uploaded_file.filename):
            continue

        image = Image.open(uploaded_file.stream).convert("RGB")
        try:
            ocr_text = ocr_image(image)
        except RuntimeError as exc:
            raise
        trades = parse_text_to_trades(ocr_text)
        if trades:
            trade_rows.extend(trades)
    return trade_rows


def parse_price(value: str) -> float | None:
    if not value:
        return None
    cleaned = value.replace(",", "")
    match = re.search(r"[-+]?[0-9]*\.?[0-9]+", cleaned)
    if not match:
        return None
    try:
        return float(match.group())
    except ValueError:
        return None


def summarize_trades(trade_rows: list[dict]) -> dict:
    total = len(trade_rows)
    wins = 0
    losses = 0
    profits = []

    outcome_counts = {"Take Profit": 0, "Stop Loss": 0, "Break Even": 0, "Manual/Other": 0}
    session_counts = {}

    for trade in trade_rows:
        profit_val = parse_price(trade.get("profit", ""))
        if profit_val is not None:
            profits.append(profit_val)
            if profit_val > 0:
                wins += 1
            elif profit_val < 0:
                losses += 1

        outcome = trade.get("outcome", "")
        if outcome in outcome_counts:
            outcome_counts[outcome] += 1

        session = trade.get("session", "")
        if session:
            session_counts[session] = session_counts.get(session, 0) + 1

    decided = wins + losses
    win_values = [p for p in profits if p > 0]
    loss_values = [p for p in profits if p < 0]

    summary = {
        "Total Trades": total,
        "Win Rate (%)": round((wins / decided * 100), 2) if decided else 0,
        "Wins": wins,
        "Losses": losses,
        "Take Profit Hits": outcome_counts["Take Profit"],
        "Stop Loss Hits": outcome_counts["Stop Loss"],
        "Break Even Exits": outcome_counts["Break Even"],
        "Manual/Other Exits": outcome_counts["Manual/Other"],
        "Net Profit": round(sum(profits), 2) if profits else 0,
        "Average Win": round(mean(win_values), 2) if win_values else "",
        "Average Loss": round(mean(loss_values), 2) if loss_values else "",
        "Session Breakdown": session_counts,
    }
    return summary


TRADE_COLUMNS = [
    "symbol", "direction", "lot_size", "entry", "exit_price",
    "stop_loss", "take_profit", "outcome", "session",
    "open_date", "open_time", "close_date", "close_time", "profit",
]
TRADE_HEADER_LABELS = [
    "Symbol", "Direction", "Lot Size", "Entry", "Exit Price",
    "Stop Loss", "Take Profit", "Outcome", "Session",
    "Open Date", "Open Time", "Close Date", "Close Time", "Profit",
]


def _trades_to_dataframe(trade_rows: list[dict]) -> pd.DataFrame:
    clean_rows = [{col: trade.get(col, "") for col in TRADE_COLUMNS} for trade in trade_rows]
    df = pd.DataFrame(clean_rows)
    if df.empty:
        df = pd.DataFrame([{col: "" for col in TRADE_COLUMNS}])
    df = df[TRADE_COLUMNS]
    df.columns = TRADE_HEADER_LABELS
    return df


def build_raw_excel_bytes(trade_rows: list[dict]) -> io.BytesIO:
    df = _trades_to_dataframe(trade_rows)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Trades")
    output.seek(0)
    return output


def build_summary_excel_bytes(trade_rows: list[dict]) -> io.BytesIO:
    summary = summarize_trades(trade_rows)
    session_breakdown = summary.pop("Session Breakdown", {})

    summary_rows = [[key, value] for key, value in summary.items()]
    summary_df = pd.DataFrame(summary_rows, columns=["Metric", "Value"])

    session_rows = [[session, count] for session, count in session_breakdown.items()]
    session_df = pd.DataFrame(session_rows, columns=["Session", "Trade Count"]) if session_rows else pd.DataFrame(
        columns=["Session", "Trade Count"]
    )

    trades_df = _trades_to_dataframe(trade_rows)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
        session_startrow = len(summary_df) + 2
        session_df.to_excel(writer, index=False, sheet_name="Summary", startrow=session_startrow)
        trades_df.to_excel(writer, index=False, sheet_name="Trades")

        summary_ws = writer.sheets["Summary"]
        chart = BarChart()
        chart.type = "col"
        chart.title = "Exit Outcome Counts"
        chart.y_axis.title = "Count"
        chart.x_axis.title = "Outcome"

        # Rows 6-9 hold Take Profit Hits, Stop Loss Hits, Break Even Exits,
        # Manual/Other Exits (row 1 is the header, row 2 is Total Trades, etc.)
        cats = Reference(summary_ws, min_col=1, min_row=6, max_row=9)
        vals = Reference(summary_ws, min_col=2, min_row=6, max_row=9)
        chart.add_data(vals, titles_from_data=False)
        chart.set_categories(cats)
        summary_ws.add_chart(chart, "D2")

        for column_cells in summary_ws.columns:
            max_length = max(
                len(str(cell.value or "")) for cell in column_cells
            )
            adjusted_width = min(40, max_length + 2)
            summary_ws.column_dimensions[column_cells[0].column_letter].width = adjusted_width

    output.seek(0)
    return output


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    if "files" not in request.files:
        flash("Please upload at least one image file.")
        return redirect(url_for("index"))

    files = request.files.getlist("files")
    trade_rows = []

    try:
        trade_rows = parse_uploaded_files(files)
    except RuntimeError as exc:
        flash(str(exc))
        return redirect(url_for("index"))

    if not trade_rows:
        flash("No valid trade data could be extracted from the uploaded screenshots.")
        return redirect(url_for("index"))

    raw_bytes = build_raw_excel_bytes(trade_rows)
    summary_bytes = build_summary_excel_bytes(trade_rows)

    bundle = io.BytesIO()
    with zipfile.ZipFile(bundle, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("trade_journal.xlsx", raw_bytes.getvalue())
        archive.writestr("trade_summary.xlsx", summary_bytes.getvalue())

    bundle.seek(0)
    return send_file(
        bundle,
        as_attachment=True,
        download_name="trade_journal_bundle.zip",
        mimetype="application/zip",
    )


@app.route("/parse", methods=["POST"])
def parse_preview():
    if "files" not in request.files:
        return jsonify({"success": False, "error": "No files uploaded."}), 400

    files = request.files.getlist("files")
    try:
        trade_rows = parse_uploaded_files(files)
    except RuntimeError as exc:
        return jsonify({"success": False, "error": str(exc)}), 400

    if not trade_rows:
        return jsonify({"success": False, "error": "No valid trade data could be extracted."}), 400

    summary = summarize_trades(trade_rows)
    return jsonify({"success": True, "trades": trade_rows, "summary": summary})


@app.route("/download", methods=["POST"])
def download_bundle():
    if "files" not in request.files:
        return jsonify({"success": False, "error": "No files uploaded."}), 400

    files = request.files.getlist("files")
    try:
        trade_rows = parse_uploaded_files(files)
    except RuntimeError as exc:
        return jsonify({"success": False, "error": str(exc)}), 400

    if not trade_rows:
        return jsonify({"success": False, "error": "No valid trade data could be extracted."}), 400

    raw_bytes = build_raw_excel_bytes(trade_rows)
    summary_bytes = build_summary_excel_bytes(trade_rows)
    bundle = io.BytesIO()
    with zipfile.ZipFile(bundle, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("trade_journal.xlsx", raw_bytes.getvalue())
        archive.writestr("trade_summary.xlsx", summary_bytes.getvalue())
    bundle.seek(0)

    return send_file(
        bundle,
        as_attachment=True,
        download_name="trade_journal_bundle.zip",
        mimetype="application/zip",
    )


if __name__ == "__main__":
    debug_mode = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=debug_mode)
