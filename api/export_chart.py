import json
import io
import base64
from http.server import BaseHTTPRequestHandler
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.styles import Font, PatternFill, Alignment, numbers
from openpyxl.utils import get_column_letter

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

HEADER_FILL  = PatternFill("solid", fgColor="374151")
STRIPE_FILL  = PatternFill("solid", fgColor="F3F4F6")
WHITE_FILL   = PatternFill("solid", fgColor="FFFFFF")
ORANGE_COLOR = "E8521A"


def build_xlsx(payload):
    division_name = payload.get("divisionName", "Division")
    labels        = payload.get("labels", [])
    values        = payload.get("displayValues", [])
    budget        = payload.get("budget", 0)
    fy_bu         = payload.get("fyBU", 0)
    upside        = payload.get("upsideCount", 0)
    gap           = payload.get("gap", 0)
    sites         = payload.get("sites", [])

    wb = Workbook()

    # ── Sheet 1: Waterfall Chart ──────────────────────────────
    ws = wb.active
    ws.title = "Waterfall Chart"
    ws.sheet_view.showGridLines = False

    # Title
    ws.merge_cells("A1:H1")
    title_cell = ws["A1"]
    title_cell.value = f"{division_name} — 2026 Pipeline Waterfall"
    title_cell.font  = Font(bold=True, size=14, color=ORANGE_COLOR)
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 24

    # Blank row
    ws.row_dimensions[2].height = 6

    # Column headers (row 3)
    for col, header in enumerate(["Category", "Value"], 1):
        cell = ws.cell(row=3, column=col, value=header)
        cell.font      = Font(bold=True, color="FFFFFF", size=10)
        cell.fill      = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 18

    # Data rows (rows 4+)
    data_start_row = 4
    for i, (lbl, val) in enumerate(zip(labels, values)):
        row_num = data_start_row + i
        fill = STRIPE_FILL if i % 2 == 0 else WHITE_FILL
        hex_col = COLOR_MAP.get(lbl, "6B7280")

        lbl_cell = ws.cell(row=row_num, column=1, value=lbl)
        lbl_cell.font      = Font(bold=(lbl in ("FY BU", "Budget")), color=hex_col)
        lbl_cell.fill      = fill
        lbl_cell.alignment = Alignment(horizontal="left", vertical="center")

        val_cell = ws.cell(row=row_num, column=2, value=val)
        val_cell.font      = Font(bold=True, color=hex_col)
        val_cell.fill      = fill
        val_cell.alignment = Alignment(horizontal="center", vertical="center")
        val_cell.number_format = "0"

        ws.row_dimensions[row_num].height = 17

    data_end_row = data_start_row + len(labels) - 1

    # Summary block
    sum_row = data_end_row + 3
    ws.cell(sum_row, 1, "Summary").font = Font(bold=True, size=11, color="374151")

    gap_color = "059669" if gap >= 0 else "DC2626"

    ws.cell(sum_row+1, 1, "FY BU vs Budget Gap").font  = Font(color="6B7280")
    ws.cell(sum_row+1, 2, gap).font                    = Font(bold=True, color=gap_color)
    ws.cell(sum_row+1, 2).number_format                = '+0;-0;0'

    ws.cell(sum_row+2, 1, "Gap %").font                = Font(color="6B7280")
    pct_cell = ws.cell(sum_row+2, 2, gap / budget if budget else 0)
    pct_cell.font          = Font(bold=True, color=gap_color)
    pct_cell.number_format = "0%"

    ws.cell(sum_row+3, 1, "FY BU + Upside").font      = Font(color="6B7280")
    ws.cell(sum_row+3, 2, fy_bu + upside).font         = Font(bold=True, color="374151")

    # Column widths
    ws.column_dimensions["A"].width = 17
    ws.column_dimensions["B"].width = 10

    # ── Embedded Bar Chart ────────────────────────────────────
    chart = BarChart()
    chart.type      = "col"
    chart.grouping  = "clustered"
    chart.overlap   = 0
    chart.title     = f"{division_name} — 2026 Pipeline"
    chart.width     = 24
    chart.height    = 15
    chart.legend    = None

    # Hide gridlines
    chart.y_axis.majorGridlines = None
    chart.y_axis.delete         = True   # hide y-axis labels, values shown as data labels
    chart.x_axis.tickLblPos     = "low"

    data_ref = Reference(ws, min_col=2, min_row=3, max_row=data_end_row)
    cats_ref = Reference(ws, min_col=1, min_row=data_start_row, max_row=data_end_row)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)

    # Per-bar colors
    ser = chart.series[0]
    ser.graphicalProperties.solidFill = "E8521A"  # default (overridden per point)
    ser.graphicalProperties.line.solidFill = "FFFFFF"

    for i, lbl in enumerate(labels):
        hex_col = COLOR_MAP.get(lbl, "9CA3AF")
        dp = DataPoint(idx=i)
        dp.graphicalProperties.solidFill        = hex_col
        dp.graphicalProperties.line.solidFill   = hex_col
        ser.dPt.append(dp)

    # Data labels above bars
    from openpyxl.chart.label import DataLabel, DataLabelList
    dl = DataLabelList()
    dl.showVal     = True
    dl.showCatName = False
    dl.showSerName = False
    dl.showPercent = False
    dl.showLegendKey = False
    dl.position    = "outEnd"
    ser.dLbls = dl

    # Place chart to the right of data (column D, row 2)
    ws.add_chart(chart, "D2")

    # ── Sheet 2: Site Detail ──────────────────────────────────
    ws2 = wb.create_sheet("Site Detail")
    ws2.sheet_view.showGridLines = False

    site_headers = [
        "SIP ID", "Rest No", "FZ", "Address", "City", "ST",
        "Status", "FZ Proj Open Date", "PLK Proj Open Date",
        "Risk Level", "Last Comments"
    ]
    col_widths = [12, 10, 20, 24, 16, 6, 12, 18, 18, 12, 44]

    for col, (h, w) in enumerate(zip(site_headers, col_widths), 1):
        cell = ws2.cell(row=1, column=col, value=h)
        cell.font      = Font(bold=True, color="FFFFFF", size=10)
        cell.fill      = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
        ws2.column_dimensions[get_column_letter(col)].width = w
    ws2.row_dimensions[1].height = 18

    risk_colors = {
        "low":    "D1FAE5", "medium": "FEF3C7", "high": "FEE2E2",
        "upside": "EDE9FE", "2027+":  "E0F2FE",
    }

    for row_i, s in enumerate(sites, 2):
        row_vals = [
            s.get("sipId",""), s.get("restNum",""), s.get("fz",""),
            s.get("address",""), s.get("city",""), s.get("state",""),
            s.get("status",""), s.get("fzOpenDate",""), s.get("plkOpenDate",""),
            s.get("riskLevel",""), s.get("lastComment","")
        ]
        for col, val in enumerate(row_vals, 1):
            cell = ws2.cell(row=row_i, column=col, value=val)
            cell.alignment = Alignment(vertical="center", wrap_text=(col == 11))
        ws2.row_dimensions[row_i].height = 16

        # Color-code risk level cell
        risk_val = s.get("riskLevel","").strip().lower()
        risk_hex = risk_colors.get(risk_val, None)
        if risk_hex:
            ws2.cell(row=row_i, column=10).fill = PatternFill("solid", fgColor=risk_hex)

    # Freeze header row
    ws2.freeze_panes = "A2"

    # Auto-filter
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(site_headers))}1"

    # ── Serialize ─────────────────────────────────────────────
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


class handler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200)
        self._set_cors()
        self.end_headers()

    def do_POST(self):
        length  = int(self.headers.get("Content-Length", 0))
        payload = json.loads(self.rfile.read(length))
        division_name = payload.get("divisionName", "Division")

        try:
            xlsx_bytes = build_xlsx(payload)
        except Exception as e:
            self.send_response(500)
            self._set_cors()
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"error": str(e)}).encode())
            return

        safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in division_name)
        filename  = f"Waterfall_{safe_name}.xlsx"

        self.send_response(200)
        self._set_cors()
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.send_header("Content-Length", str(len(xlsx_bytes)))
        self.end_headers()
        self.wfile.write(xlsx_bytes)

    def _set_cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
