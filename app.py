"""
Skat Positivliste – Flask backend
Run with: python app.py
Then open: http://localhost:5000
"""

import io
from datetime import datetime
import requests
import openpyxl
from flask import Flask, jsonify, send_from_directory, request

app = Flask(__name__, static_folder=".")

EXCEL_URL = "https://skat.dk/media/btpf4wfr/februar-2026-abis-liste-2021-2026.xlsx"

_cache = {"rows": None, "headers": None}


def load_data():
    """Download and parse the Excel file, caching the result."""
    if _cache["rows"] is not None:
        return _cache["headers"], _cache["rows"]

    print("Downloading Excel from Skat.dk…")
    resp = requests.get(EXCEL_URL, timeout=30)
    resp.raise_for_status()
    raw_bytes = resp.content

    wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), read_only=True, data_only=True)
    all_sheets = wb.sheetnames
    print(f"Sheets in workbook: {all_sheets}")

    # Target the sheet named after the current year (e.g. "2026").
    # Fallback: pick the highest numeric sheet name available.
    current_year = str(datetime.now().year)
    if current_year in all_sheets:
        sheet_name = current_year
    else:
        year_sheets = [s for s in all_sheets if s.strip().isdigit()]
        sheet_name = max(year_sheets, key=int) if year_sheets else all_sheets[0]

    print(f"Using sheet: '{sheet_name}'")
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    # Debug: print first 10 rows so problems are easy to spot in the terminal
    for i, row in enumerate(rows[:10]):
        print(f"  Row {i}: {[str(c) for c in row if c is not None]}")

    # Find the header row — first row (within first 30) where a cell contains "ISIN"
    header_idx = 0
    for i, row in enumerate(rows[:30]):
        if any("ISIN" in str(c).upper() for c in row if c is not None):
            print(f"  → ISIN found in row {i}")
            header_idx = i
            break
    else:
        print("WARNING: ISIN not found in first 30 rows — defaulting to row 0")

    headers = [str(c).strip() if c is not None else "" for c in rows[header_idx]]
    print(f"Headers: {headers}")

    data_rows = [
        [str(c).strip() if c is not None else "" for c in row]
        for row in rows[header_idx + 1:]
        if any(c is not None and str(c).strip() != "" for c in row)
    ]

    _cache["headers"] = headers
    _cache["rows"] = data_rows
    print(f"Loaded {len(data_rows)} data rows.\n")
    return headers, data_rows


@app.route("/")
def index():
    return send_from_directory(".", "index.html")


@app.route("/api/search")
def search():
    isin = request.args.get("isin", "").strip().upper()
    if not isin:
        return jsonify({"error": "Missing ISIN parameter"}), 400

    try:
        headers, rows = load_data()
    except Exception as e:
        return jsonify({"error": f"Could not load data: {e}"}), 502

    # Find ISIN column — match any header that contains "ISIN"
    isin_col = next(
        (i for i, h in enumerate(headers) if "ISIN" in h.upper()), None
    )
    if isin_col is None:
        return jsonify({
            "error": f"ISIN column not found. Headers were: {headers}"
        }), 500

    matches = [row for row in rows if row[isin_col].upper() == isin]

    return jsonify({
        "isin": isin,
        "headers": headers,
        "matches": matches,
        "total_rows": len(rows),
    })


@app.route("/api/reload", methods=["POST"])
def reload_cache():
    """Force a fresh download — useful when Skat publishes an updated list."""
    _cache["rows"] = None
    _cache["headers"] = None
    try:
        headers, rows = load_data()
        return jsonify({"ok": True, "rows": len(rows)})
    except Exception as e:
        return jsonify({"error": str(e)}), 502


if __name__ == "__main__":
    try:
        load_data()
    except Exception as e:
        print(f"Warning: could not pre-load data – {e}")
    app.run(debug=True, port=5000)
