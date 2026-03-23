from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import openpyxl
import io
import os
from datetime import datetime

app = Flask(__name__)
CORS(app)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "CDF_Template.xlsx")

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})

@app.route("/fill-cdf", methods=["POST"])
def fill_cdf():
    try:
        data = request.get_json()

        name          = data.get("name", "")
        email         = data.get("email", "")
        phone         = data.get("phone", "")
        location      = data.get("location", "")
        req_num       = data.get("req_num", "")
        date_submitted = data.get("date_submitted", "")
        currency      = data.get("currency", "USD")
        items         = data.get("items", [])   # [{description, speedkey, qty, unitPrice, date}]

        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        ws = wb["Cash Disbursement"]

        # ── Header ──────────────────────────────────────────────
        ws["L1"] = req_num
        ws["L2"] = date_submitted
        ws["L3"] = date_submitted
        ws["C4"] = name
        ws["C5"] = phone
        ws["I5"] = location
        ws["C6"] = email
        ws["H6"] = currency
        ws["I6"] = "X" if currency == "USD" else ""
        ws["J6"] = "X" if currency == "CDF" else ""

        # ── Line items ───────────────────────────────────────────
        start_row = 22
        grand_total = 0.0

        for i, item in enumerate(items):
            if i >= 19:
                break
            row = start_row + i
            desc       = item.get("description", "")
            item_date  = item.get("date", "")
            speedkey   = item.get("speedkey", "")
            qty        = float(item.get("qty", 1))
            unit_price = float(item.get("unitPrice", 0))
            line_total = round(qty * unit_price, 2)
            grand_total += line_total

            if item_date:
                desc = f"{desc} — {item_date}"

            ws[f"B{row}"] = desc
            ws[f"D{row}"] = speedkey
            ws[f"G{row}"] = qty
            ws[f"H{row}"] = unit_price
            ws[f"J{row}"] = line_total
            ws[f"L{row}"] = line_total

        grand_total = round(grand_total, 2)

        # ── Totals ───────────────────────────────────────────────
        ws["J41"] = grand_total
        ws["L41"] = grand_total

        # ── Receipt / clearance ──────────────────────────────────
        ws["A46"] = name
        ws["C64"] = name

        # ── Reconciliation ───────────────────────────────────────
        ws["D52"] = grand_total
        ws["D54"] = grand_total
        ws["D56"] = 0
        ws["D58"] = 0

        # ── Save to buffer & return ──────────────────────────────
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        safe_name = name.replace(" ", "_") or "Staff"
        date_str  = date_submitted.replace("/", "").replace("-", "") or "nodate"
        filename  = f"CDF_{safe_name}_{currency}_{date_str}.xlsx"

        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
