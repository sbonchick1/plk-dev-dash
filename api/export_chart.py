import json
import io
from http.server import BaseHTTPRequestHandler
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# Colour map keyed on label — covers every possible X-axis category
COLOR_MAP = {
    "Prospect":   "E8521A",
    "SA":         "E8521A",
    "PC":         "E8521A",
    "Permitting": "E8521A",
    "UC":         "E8521A",
    "Open":       "E8521A",
    "FY BU":      "7B2D8B",
    "Upside":     "C1272D",
    "Budget":     "00A99D",
}

HEADER_FILL = PatternFill("solid", fgColor="374151")
STRIPE_FILL = PatternFill("solid", fgColor="F3F4F6")
WHITE_FILL  = PatternFill("solid", fgColor="FFFFFF")


def build_xlsx(payload):
    division_name = payload.get("divisionName", "Division")
    # labels = ["Prospect","SA","PC","Permitting","UC","Open","FY BU","Upside","Budget"]
    labels        = payload.get("labels", [])
    values        = payload.get("displayValues", [])
    budget        = payload.get("budget", 0)
    fy_bu         = payload.get("fyBU", 0)
    upside        = payload.get("upsideCount", 0)
    gap           = payload.get("gap", 0)
    sites         = payload.get("sites", [])

    wb = Workbook()

    # ── Sheet 1: Waterfall Chart ──────────────────────────────────────────────
    ws = wb.active
    ws.title = "Waterfall Chart"
    ws.sheet_view.showGridLines = False

    # Title row
    ws.merge_cells("A1:H1")
    t = ws["A1"]
    t.value     = f"{division_name} — 2026 Pipeline Waterfall"
    t.font      = Font(bold=True, size=14, color="E8521A")
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 6  # spacer

    # Header row for data table
    for col, hdr in enumerate(["Category", "Value"], 1):
        c = ws.cell(row=3, column=col, value=hdr)
        c.font      = Font(bold=True, color="FFFFFF", size=10)
        c.fill      = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 18

    # Data rows — one per label/value pair
    data_start = 4
    for i, (lbl, val) in enumerate(zip(labels, values)):
        row  = data_start + i
        fill = STRIPE_FILL if i % 2 == 0 else WHITE_FILL
        hc   = COLOR_MAP.get(lbl, "6B7280")

        c1 = ws.cell(row=row, column=1, value=lbl)
        c1.font      = Font(bold=(lbl in ("FY BU", "Budget")), color=hc)
        c1.fill      = fill
        c1.alignment = Alignment(horizontal="left", vertical="center")

        c2 = ws.cell(row=row, column=2, value=val)
        c2.font          = Font(bold=True, color=hc)
        c2.fill          = fill
        c2.alignment     = Alignment(horizontal="center", vertical="center")
        c2.number_format = "0"
        ws.row_dimensions[row].height = 17

    data_end = data_start + len(labels) - 1

    # Summary block below the data table
    sr = data_end + 3
    ws.cell(sr,   1, "Summary").font = Font(bold=True, size=11, color="374151")
    gc = "059669" if gap >= 0 else "DC2626"

    ws.cell(sr+1, 1, "FY BU vs Budget Gap").font = Font(color="6B7280")
    gv = ws.cell(sr+1, 2, gap)
    gv.font = Font(bold=True, color=gc)
    gv.number_format = "+0;-0;0"

    ws.cell(sr+2, 1, "Gap %").font = Font(color="6B7280")
    pv = ws.cell(sr+2, 2, gap / budget if budget else 0)
    pv.font = Font(bold=True, color=gc)
    pv.number_format = "0%"

    ws.cell(sr+3, 1, "FY BU + Upside").font = Font(color="6B7280")
    ws.cell(sr+3, 2, fy_bu + upside).font    = Font(bold=True, color="374151")

    ws.column_dimensions["A"].width = 17
    ws.column_dimensions["B"].width = 10

    # ── Embedded bar chart ────────────────────────────────────────────────────
    # Uses the labels column (col A, rows data_start:data_end) as X-axis categories
    # and the values column (col B) as bar heights — so the chart always shows
    # "Prospect, SA, PC, Permitting, UC, Open, FY BU, Upside, Budget" on the axis.
    chart = BarChart()
    chart.type     = "col"
    chart.grouping = "clustered"
    chart.overlap  = 0
    chart.title    = f"{division_name} — 2026 Pipeline"
    chart.width    = 26
    chart.height   = 15
    chart.legend   = None
    chart.y_axis.majorGridlines = None
    chart.y_axis.delete         = True
    chart.x_axis.tickLblPos     = "low"

    # Values reference: header row (row 3) + data rows
    chart.add_data(
        Reference(ws, min_col=2, min_row=3, max_row=data_end),
        titles_from_data=True
    )
    # Categories reference: label column, data rows only (no header)
    chart.set_categories(
        Reference(ws, min_col=1, min_row=data_start, max_row=data_end)
    )

    # Colour every bar individually to match the web chart
    ser = chart.series[0]
    ser.graphicalProperties.solidFill      = "E8521A"
    ser.graphicalProperties.line.solidFill = "FFFFFF"
    for i, lbl in enumerate(labels):
        dp = DataPoint(idx=i)
        hc = COLOR_MAP.get(lbl, "9CA3AF")
        dp.graphicalProperties.solidFill      = hc
        dp.graphicalProperties.line.solidFill = hc
        ser.dPt.append(dp)

    # Data labels above each bar
    dl = DataLabelList()
    dl.showVal     = True
    dl.showCatName = False
    dl.showSerName = False
    dl.showPercent = False
    dl.showLegendKey = False
    dl.position    = "outEnd"
    ser.dLbls = dl

    ws.add_chart(chart, "D2")

    # ── Sheet 2: Site Detail ──────────────────────────────────────────────────
    ws2 = wb.create_sheet("Site Detail")
    ws2.sheet_view.showGridLines = False

    hdrs   = ["SIP ID","Rest No","FZ","Address","City","ST","Status",
              "FZ Proj Open Date","PLK Proj Open Date","Risk Level","Last Comments"]
    widths = [12, 10, 20, 24, 16, 6, 12, 18, 18, 12, 44]

    for col, (h, w) in enumerate(zip(hdrs, widths), 1):
        c = ws2.cell(row=1, column=col, value=h)
        c.font      = Font(bold=True, color="FFFFFF", size=10)
        c.fill      = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws2.column_dimensions[get_column_letter(col)].width = w
    ws2.row_dimensions[1].height = 18

    risk_fills = {
        "low":    "D1FAE5",
        "medium": "FEF3C7",
        "high":   "FEE2E2",
        "upside": "EDE9FE",
        "2027+":  "E0F2FE",
    }

    for ri, s in enumerate(sites, 2):
        vals = [
            s.get("sipId",""),   s.get("restNum",""), s.get("fz",""),
            s.get("address",""), s.get("city",""),     s.get("state",""),
            s.get("status",""),  s.get("fzOpenDate",""), s.get("plkOpenDate",""),
            s.get("riskLevel",""), s.get("lastComment",""),
        ]
        for col, val in enumerate(vals, 1):
            cell = ws2.cell(row=ri, column=col, value=val)
            cell.alignment = Alignment(vertical="center", wrap_text=(col == 11))
        ws2.row_dimensions[ri].height = 16

        rf = risk_fills.get(s.get("riskLevel", "").strip().lower())
        if rf:
            ws2.cell(row=ri, column=10).fill = PatternFill("solid", fgColor=rf)

    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(hdrs))}1"

    buf = io.BytesIO()
    wb.save(buf)
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
            self._respond_error(400, f"Bad request: {e}")
            return
        try:
            xlsx_bytes = build_xlsx(payload)
        except Exception as e:
            self._respond_error(500, str(e))
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

    def _respond_error(self, code, msg):
        body = json.dumps({"error": msg}).encode()
        self.send_response(code)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):
        pass
