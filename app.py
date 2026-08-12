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
        return pytesseract.image_to_string(image, lang="eng")
    except TesseractNotFoundError as exc:
        raise RuntimeError(
            "Tesseract OCR executable not found. Please install Tesseract and add it to your PATH. "
            "See README for installation instructions."
        ) from exc


def parse_ocr_text_simple(text: str) -> list[dict]:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    entries = []
    current = {"entry": "", "stop_loss": "", "take_profit": "", "break_even": ""}

    for line in lines:
        low = line.lower()
        if "entry" in low:
            current["entry"] = line.split(":", 1)[-1].strip() if ":" in line else line
        elif "stop" in low and "loss" in low:
            current["stop_loss"] = line.split(":", 1)[-1].strip() if ":" in line else line
        elif "take" in low and "profit" in low:
            current["take_profit"] = line.split(":", 1)[-1].strip() if ":" in line else line
        elif "break even" in low or "breakeven" in low:
            if "yes" in low or "true" in low or "hit" in low or "stopped" in low:
                current["break_even"] = "Yes"
            else:
                current["break_even"] = line.split(":", 1)[-1].strip() if ":" in line else "Yes"

    if any(current.values()):
        entries.append(current)
    return entries


def parse_text_to_trades(text: str) -> list[dict]:
    if openai_client:
        prompt = (
            "Extract structured trading journal entries from the following screenshot text. "
            "Return a JSON array of objects with exactly these fields: entry, stop_loss, take_profit, break_even. "
            "If the trade was stopped out by break even, set break_even to 'Yes'. Otherwise set 'No' or leave blank. "
            "If a field is missing, return an empty string. Do not include any extra keys. "
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
    break_evens = 0
    win_values = []
    loss_values = []

    for trade in trade_rows:
        be_flag = str(trade.get("break_even", "")).strip().lower()
        if be_flag == "yes":
            break_evens += 1
            continue

        if trade.get("take_profit"):
            wins += 1
        if trade.get("stop_loss"):
            losses += 1

        entry_price = parse_price(trade.get("entry", ""))
        take_price = parse_price(trade.get("take_profit", ""))
        stop_price = parse_price(trade.get("stop_loss", ""))

        if entry_price is not None and take_price is not None:
            win_values.append(abs(take_price - entry_price))
        if entry_price is not None and stop_price is not None:
            loss_values.append(abs(entry_price - stop_price))

    summary = {
        "Total Trades": total,
        "Win Rate (%)": round((wins / total * 100), 2) if total else 0,
        "Wins": wins,
        "Losses": losses,
        "Break Even": break_evens,
        "Average Win": round(mean(win_values), 4) if win_values else "",
        "Average Loss": round(mean(loss_values), 4) if loss_values else "",
    }
    return summary


def build_raw_excel_bytes(trade_rows: list[dict]) -> io.BytesIO:
    clean_rows = [
        {
            "entry": trade.get("entry", ""),
            "stop_loss": trade.get("stop_loss", ""),
            "take_profit": trade.get("take_profit", ""),
            "break_even": trade.get("break_even", ""),
        }
        for trade in trade_rows
    ]

    df = pd.DataFrame(clean_rows)
    if df.empty:
        df = pd.DataFrame(
            [{"entry": "", "stop_loss": "", "take_profit": "", "break_even": ""}]
        )

    df = df[["entry", "stop_loss", "take_profit", "break_even"]]
    df.columns = ["Entry", "Stop Loss", "Take Profit", "Break Even"]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Trades")
    output.seek(0)
    return output


def build_summary_excel_bytes(trade_rows: list[dict]) -> io.BytesIO:
    clean_rows = [
        {
            "entry": trade.get("entry", ""),
            "stop_loss": trade.get("stop_loss", ""),
            "take_profit": trade.get("take_profit", ""),
            "break_even": trade.get("break_even", ""),
        }
        for trade in trade_rows
    ]

    summary = summarize_trades(clean_rows)
    summary_rows = [
        ["Total Trades", summary["Total Trades"]],
        ["Win Rate (%)", summary["Win Rate (%)"]],
        ["Wins", summary["Wins"]],
        ["Losses", summary["Losses"]],
        ["Break Even", summary["Break Even"]],
        ["Average Win", summary["Average Win"]],
        ["Average Loss", summary["Average Loss"]],
    ]
    summary_df = pd.DataFrame(summary_rows, columns=["Metric", "Value"])
    trades_df = pd.DataFrame(clean_rows)
    trades_df = trades_df[["entry", "stop_loss", "take_profit", "break_even"]]
    trades_df.columns = ["Entry", "Stop Loss", "Take Profit", "Break Even"]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
        trades_df.to_excel(writer, index=False, sheet_name="Trades")

        summary_ws = writer.sheets["Summary"]
        chart = BarChart()
        chart.type = "col"
        chart.title = "Trade Outcome Counts"
        chart.y_axis.title = "Count"
        chart.x_axis.title = "Outcome"

        cats = Reference(summary_ws, min_col=1, min_row=3, max_row=5)
        vals = Reference(summary_ws, min_col=2, min_row=3, max_row=5)
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
