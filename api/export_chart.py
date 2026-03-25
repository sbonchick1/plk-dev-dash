import json
import io
from http.server import BaseHTTPRequestHandler
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# Colors matching the HTML page
ORANGE  = "E8521A"  # status stages
PURPLE  = "7B2D8B"  # FY BU
RED     = "C1272D"  # Upside
TEAL    = "00A99D"  # Budget
NO_FILL = "00FFFFFF"  # fully transparent (invisible base series)

HEADER_FILL = PatternFill("solid", fgColor="374151")
STRIPE_FILL = PatternFill("solid", fgColor="F3F4F6")
WHITE_FILL  = PatternFill("solid", fgColor="FFFFFF")

STAGE_LABELS  = ["Prospect", "SA", "PC", "Permitting", "UC", "Open"]
SUMMARY_LABELS = ["FY BU", "Upside", "Budget"]
ALL_LABELS    = STAGE_LABELS + SUMMARY_LABELS

# Per-label visible color
BAR_COLOR = {
    "Prospect":  ORANGE, "SA": ORANGE, "PC": ORANGE,
    "Permitting": ORANGE, "UC": ORANGE, "Open": ORANGE,
    "FY BU":  PURPLE,
    "Upside": RED,
    "Budget": TEAL,
}


def build_waterfall_rows(display_values, fy_bu, upside_count, budget):
    """
    Returns list of (label, base, visible_value) triples.
    - Status bars: base = cumulative sum of prior stages (waterfall stacking)
    - FY BU:  base = 0 (standalone)
    - Upside: base = fyBU (stacks on top of FY BU)
    - Budget: base = 0 (standalone)
    """
    rows = []
    cumulative = 0
    stage_vals = display_values[:6]

    for lbl, val in zip(STAGE_LABELS, stage_vals):
        rows.append((lbl, cumulative, val))
        cumulative += val

    rows.append(("FY BU",  0,       fy_bu))
    rows.append(("Upside", fy_bu,   upside_count))
    rows.append(("Budget", 0,       budget))
    return rows


def build_xlsx(payload):
    division_name = payload.get("divisionName", "Division")
    display_values = payload.get("displayValues", [0] * 9)
    budget       = payload.get("budget", 0)
    fy_bu        = payload.get("fyBU", 0)
    upside_count = payload.get("upsideCount", 0)
    gap          = payload.get("gap", 0)
    sites        = payload.get("sites", [])

    wb = Workbook()
    ws = wb.active
    ws.title = "Waterfall Chart"
    ws.sheet_view.showGridLines = False

    # ── Title ────────────────────────────────────────────────
    ws.merge_cells("A1:E1")
    tc = ws["A1"]
    tc.value = f"{division_name} — 2026 Pipeline Waterfall"
    tc.font  = Font(bold=True, size=14, color=ORANGE)
    tc.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 6  # spacer

    # ── Column headers (row 3) ───────────────────────────────
    headers = ["Stage", "Base (hidden)", "# SIPs"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=3, column=col, value=h)
        c.font      = Font(bold=True, color="FFFFFF", size=10)
        c.fill      = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 18

    # ── Data rows ────────────────────────────────────────────
    wf_rows = build_waterfall_rows(display_values, fy_bu, upside_count, budget)
    data_start = 4
    for i, (lbl, base, val) in enumerate(wf_rows):
        row = data_start + i
        fill = STRIPE_FILL if i % 2 == 0 else WHITE_FILL
        color = BAR_COLOR.get(lbl, "6B7280")
        is_bold = lbl in ("FY BU", "Budget")

        lc = ws.cell(row=row, column=1, value=lbl)
        lc.font = Font(bold=is_bold, color=color); lc.fill = fill
        lc.alignment = Alignment(horizontal="left", vertical="center")

        bc = ws.cell(row=row, column=2, value=base)
        bc.font = Font(color="9CA3AF"); bc.fill = fill
        bc.alignment = Alignment(horizontal="center", vertical="center")
        bc.number_format = "0"

        vc = ws.cell(row=row, column=3, value=val)
        vc.font = Font(bold=True, color=color); vc.fill = fill
        vc.alignment = Alignment(horizontal="center", vertical="center")
        vc.number_format = "0"

        ws.row_dimensions[row].height = 18

    data_end = data_start + len(wf_rows) - 1

    # ── Summary block ────────────────────────────────────────
    sum_row = data_end + 3
    gap_color = "059669" if gap >= 0 else "DC2626"

    ws.cell(sum_row, 1, "Summary").font = Font(bold=True, size=11, color="374151")

    ws.cell(sum_row+1, 1, "FY BU vs Budget Gap").font = Font(color="6B7280")
    gv = ws.cell(sum_row+1, 2, gap)
    gv.font = Font(bold=True, color=gap_color); gv.number_format = "+0;-0;0"

    ws.cell(sum_row+2, 1, "Gap %").font = Font(color="6B7280")
    pv = ws.cell(sum_row+2, 2, fy_bu / budget if budget else 0)
    pv.font = Font(bold=True, color=gap_color); pv.number_format = "0%"

    ws.cell(sum_row+3, 1, "FY BU + Upside").font = Font(color="6B7280")
    ws.cell(sum_row+3, 2, fy_bu + upside_count).font = Font(bold=True, color="374151")

    # ── Column widths ────────────────────────────────────────
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 10

    # ── Stacked Bar Chart (waterfall) ────────────────────────
    chart = BarChart()
    chart.type     = "col"
    chart.grouping = "stacked"
    chart.overlap  = 100
    chart.title    = f"{division_name} — 2026 Pipeline"
    chart.width    = 26
    chart.height   = 16
    chart.legend   = None

    chart.y_axis.majorGridlines = None
    chart.y_axis.delete         = True
    chart.x_axis.tickLblPos     = "low"

    # Series 1: invisible base
    base_ref = Reference(ws, min_col=2, min_row=data_start, max_row=data_end)
    chart.add_data(base_ref)
    base_ser = chart.series[0]
    base_ser.title = None
    base_ser.graphicalProperties.solidFill        = NO_FILL
    base_ser.graphicalProperties.line.solidFill   = NO_FILL
    base_ser.graphicalProperties.line.noFill      = True

    # Series 2: visible bars
    val_ref = Reference(ws, min_col=3, min_row=data_start, max_row=data_end)
    chart.add_data(val_ref)
    val_ser = chart.series[1]
    val_ser.title = None
    val_ser.graphicalProperties.solidFill       = ORANGE
    val_ser.graphicalProperties.line.solidFill  = "FFFFFF"

    # Per-bar colors
    for i, (lbl, _, _) in enumerate(wf_rows):
        color = BAR_COLOR.get(lbl, "9CA3AF")
        dp = DataPoint(idx=i)
        dp.graphicalProperties.solidFill       = color
        dp.graphicalProperties.line.solidFill  = "FFFFFF"
        val_ser.dPt.append(dp)

    # Data labels above bars
    from openpyxl.chart.label import DataLabelList
    dl = DataLabelList()
    dl.showVal = True; dl.showCatName = False; dl.showSerName = False
    dl.showPercent = False; dl.showLegendKey = False; dl.position = "outEnd"
    val_ser.dLbls = dl

    # Category labels from column A
    cats_ref = Reference(ws, min_col=1, min_row=data_start, max_row=data_end)
    chart.set_categories(cats_ref)

    ws.add_chart(chart, "E2")

    # ── Sheet 2: Site Detail ─────────────────────────────────
    ws2 = wb.create_sheet("Site Detail")
    ws2.sheet_view.showGridLines = False

    site_headers = ["SIP ID","Rest No","FZ","Address","City","ST",
                    "Status","FZ Proj Open Date","PLK Proj Open Date",
                    "Risk Level","Last Comments"]
    col_widths   = [12,10,20,24,16,6,12,18,18,12,44]

    for col, (h, w) in enumerate(zip(site_headers, col_widths), 1):
        c = ws2.cell(row=1, column=col, value=h)
        c.font = Font(bold=True, color="FFFFFF", size=10); c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws2.column_dimensions[get_column_letter(col)].width = w
    ws2.row_dimensions[1].height = 18

    risk_colors = {
        "low": "D1FAE5", "medium": "FEF3C7", "high": "FEE2E2",
        "upside": "EDE9FE", "2027+": "E0F2FE",
    }

    for ri, s in enumerate(sites, 2):
        vals = [s.get("sipId",""), s.get("restNum",""), s.get("fz",""),
                s.get("address",""), s.get("city",""), s.get("state",""),
                s.get("status",""), s.get("fzOpenDate",""), s.get("plkOpenDate",""),
                s.get("riskLevel",""), s.get("lastComment","")]
        for col, v in enumerate(vals, 1):
            c = ws2.cell(row=ri, column=col, value=v)
            c.alignment = Alignment(vertical="center", wrap_text=(col == 11))
        ws2.row_dimensions[ri].height = 16

        rl = s.get("riskLevel","").strip().lower()
        rh = risk_colors.get(rl)
        if rh:
            ws2.cell(row=ri, column=10).fill = PatternFill("solid", fgColor=rh)

    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(site_headers))}1"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


class handler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200)
        self._cors()
        self.end_headers()

    def do_POST(self):
        length  = int(self.headers.get("Content-Length", 0))
        payload = json.loads(self.rfile.read(length))

        try:
            xlsx_bytes = build_xlsx(payload)
        except Exception as e:
            self.send_response(500)
            self._cors()
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"error": str(e)}).encode())
            return

        safe = "".join(c if c.isalnum() or c in " _-" else "_"
                       for c in payload.get("divisionName", "Division"))
        fname = f"Waterfall_{safe}.xlsx"

        self.send_response(200)
        self._cors()
        self.send_header("Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
        self.send_header("Content-Length", str(len(xlsx_bytes)))
        self.end_headers()
        self.wfile.write(xlsx_bytes)

    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def log_message(self, fmt, *args):
        pass
