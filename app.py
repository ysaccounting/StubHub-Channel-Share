"""
StubHub Channel Share Report — backend.

Upload one or more TicketVault Invoice Details Report exports (.xlsx).
The app combines them, deduplicates, maps companies to seller accounts,
and produces two formatted Excel files — one for ystickets, one for yitzknopf.
"""

import io
import os
import re
import time
import uuid
import shutil
import tempfile

from flask import Flask, request, jsonify, send_file, send_from_directory, abort
from processor import process_sales, write_excel

app = Flask(__name__, static_folder=None)
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
STORE_DIR  = os.path.join(tempfile.gettempdir(), "ticketvault_store")
os.makedirs(STORE_DIR, exist_ok=True)


def _cleanup_old(max_age_seconds=12 * 3600):
    now = time.time()
    for name in os.listdir(STORE_DIR):
        path = os.path.join(STORE_DIR, name)
        try:
            if os.path.isdir(path) and now - os.path.getmtime(path) > max_age_seconds:
                shutil.rmtree(path, ignore_errors=True)
        except OSError:
            pass


def _safe(s):
    return re.sub(r'[\\/:*?"<>|]+', " ", s).strip() if s else s


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory(BASE_DIR, "index.html")


@app.route("/process", methods=["POST"])
def process():
    uploaded = request.files.getlist("sales_files")
    files    = [(f.filename, f.read()) for f in uploaded if f.filename]

    if not files:
        return jsonify({"error": "Please upload at least one TicketVault export."}), 400

    try:
        # Wrap raw bytes back into file-like objects for processor
        file_likes = [io.BytesIO(data) for _, data in files]
        result     = process_sales(file_likes)
    except Exception as exc:
        return jsonify({"error": f"{type(exc).__name__}: {exc}"}), 500

    month_label    = result["month_label"]
    ys_rows        = result["ys_rows"]
    yitz_rows      = result["yitz_rows"]
    unmapped       = result["unmapped"]
    total_raw      = result["total_raw"]
    total_deduped  = result["total_deduped"]
    total_filtered = result["total_filtered"]

    # Generate Excel files
    ys_buf   = io.BytesIO()
    yitz_buf = io.BytesIO()
    write_excel(ys_rows,   ys_buf)
    write_excel(yitz_rows, yitz_buf)

    # Store to temp folder for download
    token  = uuid.uuid4().hex
    folder = os.path.join(STORE_DIR, token)
    os.makedirs(folder, exist_ok=True)

    ys_filename   = f"{month_label} - ystickets.xlsx"
    yitz_filename = f"{month_label} - yitzknopf.xlsx"

    with open(os.path.join(folder, ys_filename), "wb") as fh:
        fh.write(ys_buf.getvalue())
    with open(os.path.join(folder, yitz_filename), "wb") as fh:
        fh.write(yitz_buf.getvalue())

    _cleanup_old()

    ys_sales   = sum(r["Ticket Sales"] for r in ys_rows)
    yitz_sales = sum(r["Ticket Sales"] for r in yitz_rows)

    return jsonify({
        "month_label":      month_label,
        "total_raw":        total_raw,
        "total_deduped":    total_deduped,
        "total_filtered":   total_filtered,
        "duplicates":       total_raw - total_deduped,
        "zero_dollar":      total_deduped - total_filtered,
        "ys_rows":          len(ys_rows),
        "ys_sales":         round(ys_sales, 2),
        "yitz_rows":        len(yitz_rows),
        "yitz_sales":       round(yitz_sales, 2),
        "unmapped":         sorted(unmapped),
        "download_ys":      f"/download/{token}?which=ystickets",
        "download_yitz":    f"/download/{token}?which=yitzknopf",
    })


@app.route("/download/<token>")
def download(token):
    folder = os.path.join(STORE_DIR, os.path.basename(token))
    if not os.path.isdir(folder):
        abort(404)

    which = request.args.get("which", "")
    files = [f for f in os.listdir(folder) if f.lower().endswith(".xlsx")]
    if not files:
        abort(404)

    if which == "ystickets":
        pick = next((f for f in files if "ystickets" in f), files[0])
    elif which == "yitzknopf":
        pick = next((f for f in files if "yitzknopf" in f), files[0])
    else:
        pick = files[0]

    return send_file(
        os.path.join(folder, pick),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=pick,
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
