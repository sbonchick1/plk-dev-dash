"""
export_chart.py  —  Vercel Python serverless handler
Generates a waterfall chart Excel workbook from pipeline JSON payload.
Uses xlsxwriter for reliable, corruption-free OOXML output.

Requirements (requirements.txt):
    xlsxwriter>=3.1.0
"""

import json
import io
from http.server import BaseHTTPRequestHandler
import xlsxwriter

STATUSES = ["Prospect", "SA", "PC", "Permitting", "UC", "Open"]

BAR_COLORS = {
    "Prospect":   "#E8521A",
    "SA":         "#E8521A",
    "PC":         "#E8521A",
    "Permitting": "#E8521A",
    "UC":         "#E8521A",
    "Open":       "#E8521A",
    "FY BU":      "#7B2D8B",
    "Upside":     "#C1272D",
    "Budget":     "#00A99D",
}

RISK_BG = {
    "low":    "#D1FAE5",
    "medium": "#FEF3C7",
    "high":   "#FEE2E2",
    "upside": "#EDE9FE",
    "2027+":  "#E0F2FE",
}

# Risk levels that belong in the Site Detail tab
SITE_DETAIL_RISKS = {"low", "medium", "high", "upside"}


def build_xlsx(payload: dict) -> bytes:
    division_name = payload.get("divisionName", "Division")
    labels        = payload.get("labels", [])
    values        = payload.get("displayValues", [])
    budget        = int(payload.get("budget", 0))
    fy_bu         = int(payload.get("fyBU", 0))
    upside_count  = int(payload.get("upsideCount", 0))
    gap           = int(payload.get("gap", 0))
    # FIX 2: filter to only low/medium/high/upside sites for the Site Detail tab
    all_sites = payload.get("sites", [])
    sites = [s for s in all_sites
             if s.get("riskLevel", "").strip().lower() in SITE_DETAIL_RISKS]

    bar_labels = STATUSES + ["FY BU", "Upside", "Budget"]
    n = len(bar_labels)

    # Per-status values in label order
    status_vals = []
    for lbl in STATUSES:
        idx = labels.index(lbl) if lbl in labels else -1
        status_vals.append(int(values[idx]) if idx >= 0 else 0)

    # ── Build waterfall base / visible-bar data ─────────────────────────────
    base_data = []
    bar_data  = []
    cumulative = 0
    for v in status_vals:
        base_data.append(cumulative)
        bar_data.append(v)
        cumulative += v

    # FY BU  — standalone, starts from 0
    base_data.append(0)
    bar_data.append(fy_bu)

    # Upside — stacks directly on top of FY BU
    base_data.append(fy_bu)
    bar_data.append(upside_count)

    # Budget — standalone, starts from 0
    base_data.append(0)
    bar_data.append(budget)

    # ── Workbook ────────────────────────────────────────────────────────────
    buf = io.BytesIO()
    wb  = xlsxwriter.Workbook(buf, {"in_memory": True})

    # ── Common formats ──────────────────────────────────────────────────────
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

    def gap_num_fmt(positive):
        color = "#059669" if positive else "#DC2626"
        return fmt(bold=True, font_color=color, align="center", num_format="+0;-0;0")

    def gap_pct_fmt_fn(positive):
        color = "#059669" if positive else "#DC2626"
        return fmt(bold=True, font_color=color, align="center", num_format="0%")

    def row_label_fmt(color, is_special, even):
        return fmt(
            bold=is_special,
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

    # ── Sheet 1: Waterfall Chart ────────────────────────────────────────────
    ws = wb.add_worksheet("Waterfall Chart")
    ws.hide_gridlines(2)
    ws.set_column("A:A", 14)
    ws.set_column("B:B", 12)
    ws.set_column("C:C", 10)

    ws.merge_range("A1:C1", f"{division_name} \u2014 2026 Pipeline Waterfall", title_fmt)
    ws.set_row(0, 28)
    ws.set_row(1, 6)

    ws.write(2, 0, "Stage",  hdr_fmt)
    ws.write(2, 1, "Base",   hdr_fmt)
    ws.write(2, 2, "# SIPs", hdr_fmt)
    ws.set_row(2, 18)

    DATA_ROW0 = 3
    for i, lbl in enumerate(bar_labels):
        r       = DATA_ROW0 + i
        even    = (i % 2 == 0)
        color   = BAR_COLORS.get(lbl, "#6B7280")
        special = lbl in ("FY BU", "Budget", "Upside")
        ws.write(r, 0, lbl,          row_label_fmt(color, special, even))
        ws.write(r, 1, base_data[i], row_base_fmt(even))
        ws.write(r, 2, bar_data[i],  row_val_fmt(color, even))
        ws.set_row(r, 17)

    DATA_ROW_LAST = DATA_ROW0 + n - 1

    sr    = DATA_ROW_LAST + 2
    g_pos = gap >= 0
    ws.write(sr,   0, "FY BU",                 stat_lbl)
    ws.write(sr,   1, fy_bu,                   stat_val)
    ws.write(sr+1, 0, "Budget",                stat_lbl)
    ws.write(sr+1, 1, budget,                  stat_val)
    ws.write(sr+2, 0, "Gap (FY BU vs Budget)", stat_lbl)
    ws.write(sr+2, 1, gap,                     gap_num_fmt(g_pos))
    ws.write(sr+3, 0, "Gap %",                 stat_lbl)
    ws.write(sr+3, 1, (gap / budget) if budget else 0, gap_pct_fmt_fn(g_pos))
    ws.write(sr+4, 0, "FY BU + Upside",        stat_lbl)
    ws.write(sr+4, 1, fy_bu + upside_count,    stat_val)

    # ── Chart ───────────────────────────────────────────────────────────────
    chart = wb.add_chart({"type": "column", "subtype": "stacked"})

    # Series 0 — invisible base
    chart.add_series({
        "name":       "_base",
        "categories": ["Waterfall Chart", DATA_ROW0, 0, DATA_ROW_LAST, 0],
        "values":     ["Waterfall Chart", DATA_ROW0, 1, DATA_ROW_LAST, 1],
        "fill":       {"none": True},
        "border":     {"none": True},
    })

    # Series 1 — visible bars with per-bar colors
    points = [
        {"fill": {"color": BAR_COLORS.get(lbl, "#9CA3AF")}, "border": {"none": True}}
        for lbl in bar_labels
    ]

    chart.add_series({
        "name":       "_bars",
        "categories": ["Waterfall Chart", DATA_ROW0, 0, DATA_ROW_LAST, 0],
        "values":     ["Waterfall Chart", DATA_ROW0, 2, DATA_ROW_LAST, 2],
        "fill":       {"color": "#E8521A"},
        "border":     {"none": True},
        "points":     points,
        # white labels centered within each bar
        "data_labels": {
            "value":    True,
            "position": "center",
            "font":     {"bold": True, "size": 10, "color": "#FFFFFF"},
        },
    })

    chart.set_title({"name": f"{division_name} \u2014 2026 Pipeline", "overlay": False})
    chart.set_legend({"none": True})
    chart.set_size({"width": 680, "height": 400})
    chart.set_y_axis({"visible": False, "major_gridlines": {"visible": False}})
    chart.set_x_axis({"major_gridlines": {"visible": False}, "line": {"none": True}})
    chart.set_chartarea({"border": {"none": True}})
    chart.set_plotarea({"border": {"none": True}})

    ws.insert_chart("E2", chart)

    # ── Sheet 2: Site Detail (low + medium + high + upside) ─────────────────
    ws2 = wb.add_worksheet("Site Detail")
    ws2.hide_gridlines(2)
    ws2.freeze_panes(1, 0)
    ws2.autofilter(0, 0, 0, 10)

    col_hdrs   = ["SIP ID", "Rest No", "FZ", "Address", "City", "ST", "Status",
                  "FZ Proj Open Date", "PLK Proj Open Date", "Risk Level", "Last Comments"]
    col_widths = [12, 10, 20, 24, 16, 6, 12, 18, 18, 12, 44]

    for col, (h, w) in enumerate(zip(col_hdrs, col_widths)):
        ws2.write(0, col, h, hdr_fmt)
        ws2.set_column(col, col, w)
    ws2.set_row(0, 18)

    for ri, s in enumerate(sites, 1):
        row_vals = [
            s.get("sipId",      ""),
            s.get("restNum",    ""),
            s.get("fz",         ""),
            s.get("address",    ""),
            s.get("city",       ""),
            s.get("state",      ""),
            s.get("status",     ""),
            s.get("fzOpenDate", ""),
            s.get("plkOpenDate",""),
            s.get("riskLevel",  ""),
            s.get("lastComment",""),
        ]
        for col, val in enumerate(row_vals):
            ws2.write(ri, col, val, wrap_fmt if col == 10 else cell_fmt)
        ws2.set_row(ri, 16)

        # Risk-level background on the Risk Level column
        rl = s.get("riskLevel", "").strip().lower()
        if rl in RISK_BG:
            risk_fmt = fmt(bg_color=RISK_BG[rl])
            ws2.write(ri, 9, s.get("riskLevel", ""), risk_fmt)

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
        filename = f"Waterfall_{safe}.xlsx"

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
