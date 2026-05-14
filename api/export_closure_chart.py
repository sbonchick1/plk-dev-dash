"""
export_closure_chart.py  —  Vercel Python serverless handler
Generates a closure reason waterfall Excel workbook with embedded stacked chart.
Uses xlsxwriter for reliable, corruption-free OOXML output.

Requirements (requirements.txt):
    xlsxwriter>=3.1.0
"""

import json
import io
from http.server import BaseHTTPRequestHandler
import xlsxwriter

RISK_BG = {
    "low risk":           "#D1FAE5",
    "medium risk":        "#FEF3C7",
    "high risk":          "#FEE2E2",
    "definite closure":   "#FEE2E2",
    "closed":             "#E5E7EB",
}


def build_xlsx(payload: dict) -> bytes:
    bucket_name    = payload.get("bucketName", "Closure LE")
    reasons        = payload.get("reasons", [])
    ctrl_counts    = [int(x) for x in payload.get("ctrlCounts", [])]
    nc_counts      = [int(x) for x in payload.get("nonCtrlCounts", [])]
    grand_total    = int(payload.get("grandTotal", 0))
    ctrl_color     = payload.get("ctrlColor", "#7B2D8B")
    non_ctrl_color = payload.get("nonCtrlColor", "#C1272D")
    total_color    = payload.get("totalColor", "#00A99D")
    sites          = payload.get("sites", [])

    # All labels: reasons + Total
    all_labels = list(reasons) + ["Total"]
    n = len(all_labels)

    # Build waterfall base / per-series bar data
    base_data  = []
    c_data     = []
    nc_data    = []
    tot_data   = []
    cumulative = 0

    for i in range(len(reasons)):
        c  = ctrl_counts[i]  if i < len(ctrl_counts) else 0
        nc = nc_counts[i]    if i < len(nc_counts)   else 0
        base_data.append(cumulative)
        c_data.append(c)
        nc_data.append(nc)
        tot_data.append(0)
        cumulative += c + nc

    # Total bar — standalone, starts from 0
    base_data.append(0)
    c_data.append(0)
    nc_data.append(0)
    tot_data.append(grand_total)

    # ── Workbook ─────────────────────────────────────────────────────────────
    buf = io.BytesIO()
    wb  = xlsxwriter.Workbook(buf, {"in_memory": True})

    def fmt(**kw):
        defaults = {"font_name": "Calibri", "font_size": 10, "valign": "vcenter"}
        defaults.update(kw)
        return wb.add_format(defaults)

    hdr_fmt   = fmt(bold=True, font_color="#FFFFFF", bg_color="#374151", align="center", border=0)
    title_fmt = fmt(bold=True, font_color="#E8521A", font_size=14)
    cell_fmt  = fmt()
    wrap_fmt  = fmt(text_wrap=True)

    def lbl_fmt(color, even):
        return fmt(font_color=color,
                   bg_color="#F3F4F6" if even else "#FFFFFF",
                   align="left")

    def num_fmt(color, even):
        return fmt(bold=True, font_color=color,
                   bg_color="#F3F4F6" if even else "#FFFFFF",
                   align="center", num_format="0")

    def base_cell_fmt(even):
        return fmt(font_color="#9CA3AF",
                   bg_color="#F3F4F6" if even else "#FFFFFF",
                   align="center", num_format="0")

    # ── Sheet 1: Waterfall Chart ──────────────────────────────────────────────
    ws = wb.add_worksheet("Waterfall Chart")
    ws.hide_gridlines(2)
    ws.set_column("A:A", 30)
    ws.set_column("B:B", 10)
    ws.set_column("C:C", 16)
    ws.set_column("D:D", 16)
    ws.set_column("E:E", 10)

    ws.merge_range("A1:E1", f"2026 Closure Reason — {bucket_name}", title_fmt)
    ws.set_row(0, 28)
    ws.set_row(1, 6)

    ws.write(2, 0, "Reason",         hdr_fmt)
    ws.write(2, 1, "Base",           hdr_fmt)
    ws.write(2, 2, "PLK Controlled", hdr_fmt)
    ws.write(2, 3, "Non-Controlled", hdr_fmt)
    ws.write(2, 4, "Total",          hdr_fmt)
    ws.set_row(2, 18)

    DATA_ROW0 = 3
    for i, lbl in enumerate(all_labels):
        r        = DATA_ROW0 + i
        even     = (i % 2 == 0)
        is_total = (i == n - 1)
        row_color = total_color if is_total else "#374151"

        ws.write(r, 0, lbl,          lbl_fmt(row_color, even))
        ws.write(r, 1, base_data[i], base_cell_fmt(even))
        ws.write(r, 2, c_data[i],    num_fmt(ctrl_color if not is_total else "#9CA3AF", even))
        ws.write(r, 3, nc_data[i],   num_fmt(non_ctrl_color if not is_total else "#9CA3AF", even))
        ws.write(r, 4, tot_data[i],  num_fmt(total_color if is_total else "#9CA3AF", even))
        ws.set_row(r, 17)

    DATA_ROW_LAST = DATA_ROW0 + n - 1

    # Summary stats below the table
    sr = DATA_ROW_LAST + 2
    stat_lbl = fmt(font_color="#6B7280", align="left")
    stat_val = fmt(bold=True, font_color="#374151", align="center", num_format="0")
    ctrl_total = sum(c_data)
    nc_total   = sum(nc_data)
    ws.write(sr,   0, "PLK Controlled",  stat_lbl)
    ws.write(sr,   1, ctrl_total,        stat_val)
    ws.write(sr+1, 0, "Non-Controlled",  stat_lbl)
    ws.write(sr+1, 1, nc_total,          stat_val)
    ws.write(sr+2, 0, "Grand Total",     stat_lbl)
    ws.write(sr+2, 1, grand_total,       stat_val)

    # ── Chart ─────────────────────────────────────────────────────────────────
    chart = wb.add_chart({"type": "column", "subtype": "stacked"})

    # Series 0 — invisible base spacer
    chart.add_series({
        "name":       "_base",
        "categories": ["Waterfall Chart", DATA_ROW0, 0, DATA_ROW_LAST, 0],
        "values":     ["Waterfall Chart", DATA_ROW0, 1, DATA_ROW_LAST, 1],
        "fill":       {"none": True},
        "border":     {"none": True},
    })

    # Series 1 — PLK Controlled (zeros suppressed via num_format)
    chart.add_series({
        "name":       "PLK Controlled",
        "categories": ["Waterfall Chart", DATA_ROW0, 0, DATA_ROW_LAST, 0],
        "values":     ["Waterfall Chart", DATA_ROW0, 2, DATA_ROW_LAST, 2],
        "fill":       {"color": ctrl_color},
        "border":     {"none": True},
        "data_labels": {
            "value":      True,
            "position":   "center",
            "font":       {"bold": True, "size": 10, "color": "#FFFFFF"},
            "num_format": '0;0;""',
        },
    })

    # Series 2 — Non-Controlled
    chart.add_series({
        "name":       "Non-Controlled",
        "categories": ["Waterfall Chart", DATA_ROW0, 0, DATA_ROW_LAST, 0],
        "values":     ["Waterfall Chart", DATA_ROW0, 3, DATA_ROW_LAST, 3],
        "fill":       {"color": non_ctrl_color},
        "border":     {"none": True},
        "data_labels": {
            "value":      True,
            "position":   "center",
            "font":       {"bold": True, "size": 10, "color": "#FFFFFF"},
            "num_format": '0;0;""',
        },
    })

    # Series 3 — Total bar
    chart.add_series({
        "name":       "Total",
        "categories": ["Waterfall Chart", DATA_ROW0, 0, DATA_ROW_LAST, 0],
        "values":     ["Waterfall Chart", DATA_ROW0, 4, DATA_ROW_LAST, 4],
        "fill":       {"color": total_color},
        "border":     {"none": True},
        "data_labels": {
            "value":      True,
            "position":   "center",
            "font":       {"bold": True, "size": 10, "color": "#FFFFFF"},
            "num_format": '0;0;""',
        },
    })

    chart.set_title({"name": f"2026 Closure Reason — {bucket_name}", "overlay": False})
    chart.set_legend({"position": "bottom", "delete_series": [0]})
    chart.set_size({"width": 720, "height": 440})
    chart.set_y_axis({"visible": False, "major_gridlines": {"visible": False}})
    chart.set_x_axis({"major_gridlines": {"visible": False}, "line": {"none": True}})
    chart.set_chartarea({"border": {"none": True}})
    chart.set_plotarea({"border": {"none": True}})

    ws.insert_chart("G2", chart)

    # ── Sheet 2: Site Detail ──────────────────────────────────────────────────
    ws2 = wb.add_worksheet("Site Detail")
    ws2.hide_gridlines(2)
    ws2.freeze_panes(1, 0)
    ws2.autofilter(0, 0, 0, 12)

    col_hdrs   = ["Rest No", "FZ", "Address", "City", "ST", "Date of Closure",
                  "Closure Bucket", "Closure Risk", "Closure Reason", "PLK Control",
                  "2025 ARS", "TTM EBITDA", "Comments"]
    col_widths = [10, 22, 24, 16, 6, 14, 14, 16, 24, 12, 12, 12, 40]

    for col, (h, w) in enumerate(zip(col_hdrs, col_widths)):
        ws2.write(0, col, h, hdr_fmt)
        ws2.set_column(col, col, w)
    ws2.set_row(0, 18)

    for ri, s in enumerate(sites, 1):
        row_vals = [
            s.get("restNum",       ""),
            s.get("fz",            ""),
            s.get("address",       ""),
            s.get("city",          ""),
            s.get("state",         ""),
            s.get("dateOfClosure", ""),
            s.get("closureBucket", ""),
            s.get("closureRisk",   ""),
            s.get("closureReason", ""),
            s.get("plkControl",    ""),
            s.get("ars2025",       ""),
            s.get("ttmEbitda",     ""),
            s.get("comments",      ""),
        ]
        for col, val in enumerate(row_vals):
            ws2.write(ri, col, val, wrap_fmt if col == 12 else cell_fmt)
        ws2.set_row(ri, 16)

        rl = s.get("closureRisk", "").strip().lower()
        if rl in RISK_BG:
            risk_fmt = fmt(bg_color=RISK_BG[rl])
            ws2.write(ri, 7, s.get("closureRisk", ""), risk_fmt)

    wb.close()
    buf.seek(0)
    return buf.read()


class handler(BaseHTTPRequestHandler):

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin",  "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_POST(self):
        try:
            length  = int(self.headers.get("Content-Length", 0))
            payload = json.loads(self.rfile.read(length))
        except Exception as e:
            self._error(400, "Bad request: " + str(e))
            return

        try:
            xlsx_bytes = build_xlsx(payload)
        except Exception as e:
            self._error(500, str(e))
            return

        bucket_name = payload.get("bucketName", "Closure")
        safe        = "".join(c if c.isalnum() or c in " _-" else "_" for c in bucket_name)
        filename    = f"Closure_Reason_{safe}.xlsx"

        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.send_header("Content-Length", str(len(xlsx_bytes)))
        self.end_headers()
        self.wfile.write(xlsx_bytes)

    def _error(self, code: int, msg: str):
        body = json.dumps({"error": msg}).encode()
        self.send_response(code)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):
        pass
