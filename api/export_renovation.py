"""
export_renovation.py  —  Vercel Python serverless handler
Generates a renovation stage waterfall Excel workbook with an embedded chart.
Matches the style of export_chart.py used by the development execution dashboard.

Requirements (requirements.txt):
    xlsxwriter>=3.1.0
"""

import json
import io
from http.server import BaseHTTPRequestHandler
import xlsxwriter

STAGES = [
    "Not Started",
    "360 Phase",
    "In Design",
    "Permitting",
    "Under Construction",
    "Punch List",
    "Complete",
    "Total LE",
]

STAGE_COLORS = {
    "Not Started":        "#E8521A",
    "360 Phase":          "#E8521A",
    "In Design":          "#E8521A",
    "Permitting":         "#E8521A",
    "Under Construction": "#E8521A",
    "Punch List":         "#E8521A",
    "Complete":           "#E8521A",
    "Total LE":           "#00A99D",
}


def build_xlsx(payload: dict) -> bytes:
    division_name = payload.get("divisionName", "Division")
    bu_filter     = payload.get("buFilter", "")
    stage_counts  = payload.get("stageCounts", {})   # {stage: count}
    total         = int(payload.get("total", 0))
    sites         = payload.get("sites", [])          # [{plk, division, bu, stage, rawStatus}]

    # ── Waterfall base / bar data ────────────────────────────────────────────
    # Stages stack cumulatively; Total LE starts from 0
    base_data = []
    bar_data  = []
    cumulative = 0
    for stage in STAGES:
        if stage == "Total LE":
            base_data.append(0)
            bar_data.append(total)
        else:
            v = int(stage_counts.get(stage, 0))
            base_data.append(cumulative)
            bar_data.append(v)
            cumulative += v

    n = len(STAGES)

    # ── Workbook ─────────────────────────────────────────────────────────────
    buf = io.BytesIO()
    wb  = xlsxwriter.Workbook(buf, {"in_memory": True})

    def fmt(**kw):
        defaults = {"font_name": "Calibri", "font_size": 10, "valign": "vcenter"}
        defaults.update(kw)
        return wb.add_format(defaults)

    hdr_fmt   = fmt(bold=True, font_color="#FFFFFF", bg_color="#374151", align="center", border=0)
    title_fmt = fmt(bold=True, font_color="#E8521A", font_size=14)
    stat_lbl  = fmt(font_color="#6B7280", align="left")
    stat_val  = fmt(bold=True, font_color="#374151", align="center", num_format="0")
    cell_fmt  = fmt()
    wrap_fmt  = fmt(text_wrap=True)

    def row_label_fmt(color, special, even):
        return fmt(
            bold=special,
            font_color=color,
            bg_color="#F3F4F6" if even else "#FFFFFF",
            align="left",
        )

    def row_val_fmt(color, even):
        return fmt(
            bold=True,
            font_color=color,
            bg_color="#F3F4F6" if even else "#FFFFFF",
            align="center",
            num_format="0",
        )

    def row_base_fmt(even):
        return fmt(
            font_color="#9CA3AF",
            bg_color="#F3F4F6" if even else "#FFFFFF",
            align="center",
            num_format="0",
        )

    # ── Sheet 1: Waterfall Chart ─────────────────────────────────────────────
    ws = wb.add_worksheet("Waterfall Chart")
    ws.hide_gridlines(2)
    ws.set_column("A:A", 20)
    ws.set_column("B:B", 10)
    ws.set_column("C:C", 10)

    subtitle = f"BU {bu_filter}" if bu_filter else "All BUs"
    ws.merge_range("A1:C1", f"{division_name} — Renovation Waterfall ({subtitle})", title_fmt)
    ws.set_row(0, 28)
    ws.set_row(1, 6)

    ws.write(2, 0, "Stage",   hdr_fmt)
    ws.write(2, 1, "Base",    hdr_fmt)
    ws.write(2, 2, "# Renos", hdr_fmt)
    ws.set_row(2, 18)

    DATA_ROW0 = 3
    for i, stage in enumerate(STAGES):
        r       = DATA_ROW0 + i
        even    = (i % 2 == 0)
        color   = STAGE_COLORS.get(stage, "#6B7280")
        special = stage == "Total LE"
        ws.write(r, 0, stage,         row_label_fmt(color, special, even))
        ws.write(r, 1, base_data[i],  row_base_fmt(even))
        ws.write(r, 2, bar_data[i],   row_val_fmt(color, even))
        ws.set_row(r, 17)

    DATA_ROW_LAST = DATA_ROW0 + n - 1

    # Summary stats below the data table
    sr = DATA_ROW_LAST + 2
    ws.write(sr,   0, "Total LE",  stat_lbl)
    ws.write(sr,   1, total,       stat_val)
    ws.write(sr+1, 0, "Complete",  stat_lbl)
    ws.write(sr+1, 1, int(stage_counts.get("Complete", 0)), stat_val)
    pct = round(stage_counts.get("Complete", 0) / total * 100) if total else 0
    ws.write(sr+2, 0, "% Complete", stat_lbl)
    ws.write(sr+2, 1, pct / 100,
             fmt(bold=True, font_color="#059669", align="center", num_format="0%"))

    # ── Embedded chart ───────────────────────────────────────────────────────
    chart = wb.add_chart({"type": "column", "subtype": "stacked"})

    # Invisible base series
    chart.add_series({
        "name":       "_base",
        "categories": ["Waterfall Chart", DATA_ROW0, 0, DATA_ROW_LAST, 0],
        "values":     ["Waterfall Chart", DATA_ROW0, 1, DATA_ROW_LAST, 1],
        "fill":       {"none": True},
        "border":     {"none": True},
    })

    # Visible bars with per-bar colors
    points = [
        {"fill": {"color": STAGE_COLORS.get(s, "#E8521A")}, "border": {"none": True}}
        for s in STAGES
    ]
    chart.add_series({
        "name":       "_bars",
        "categories": ["Waterfall Chart", DATA_ROW0, 0, DATA_ROW_LAST, 0],
        "values":     ["Waterfall Chart", DATA_ROW0, 2, DATA_ROW_LAST, 2],
        "fill":       {"color": "#E8521A"},
        "border":     {"none": True},
        "points":     points,
        "data_labels": {
            "value":    True,
            "position": "center",
            "font":     {"bold": True, "size": 10, "color": "#FFFFFF"},
        },
    })

    chart.set_title({"name": f"{division_name} — Renovation Stage Waterfall", "overlay": False})
    chart.set_legend({"none": True})
    chart.set_size({"width": 700, "height": 420})
    chart.set_y_axis({"visible": False, "major_gridlines": {"visible": False}})
    chart.set_x_axis({"major_gridlines": {"visible": False}, "line": {"none": True}})
    chart.set_chartarea({"border": {"none": True}})
    chart.set_plotarea({"border": {"none": True}})

    ws.insert_chart("E2", chart)

    # ── Sheet 2: Restaurant Detail ───────────────────────────────────────────
    ws2 = wb.add_worksheet("Restaurant Detail")
    ws2.hide_gridlines(2)
    ws2.freeze_panes(1, 0)
    ws2.autofilter(0, 0, 0, 4)

    col_hdrs   = ["Store #", "Division", "BU", "Stage", "Raw Status"]
    col_widths = [10, 14, 8, 22, 38]

    for col, (h, w) in enumerate(zip(col_hdrs, col_widths)):
        ws2.write(0, col, h, hdr_fmt)
        ws2.set_column(col, col, w)
    ws2.set_row(0, 18)

    for ri, s in enumerate(sites, 1):
        ws2.write(ri, 0, s.get("plk",       ""), cell_fmt)
        ws2.write(ri, 1, s.get("division",  ""), cell_fmt)
        ws2.write(ri, 2, s.get("bu",        ""), cell_fmt)
        ws2.write(ri, 3, s.get("stage",     ""), cell_fmt)
        ws2.write(ri, 4, s.get("rawStatus", ""), wrap_fmt)
        ws2.set_row(ri, 16)

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

        division_name = payload.get("divisionName", "Division")
        safe     = "".join(c if c.isalnum() or c in " _-" else "_" for c in division_name)
        bu_filter = payload.get("buFilter", "")
        filename = f"Renovations_{safe}{'_BU' + bu_filter if bu_filter else ''}.xlsx"

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
