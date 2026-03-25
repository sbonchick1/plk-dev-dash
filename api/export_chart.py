import json
import io
import zipfile
import re
from http.server import BaseHTTPRequestHandler
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment
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

HEADER_FILL = PatternFill("solid", fgColor="374151")
STRIPE_FILL = PatternFill("solid", fgColor="F3F4F6")
WHITE_FILL  = PatternFill("solid", fgColor="FFFFFF")


def build_xlsx(payload):
    division_name = payload.get("divisionName", "Division")
    labels        = payload.get("labels", [])        # ["Prospect","SA",…,"FY BU","Upside","Budget"]
    values        = payload.get("displayValues", []) # matching numeric values
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

    # Title
    ws.merge_cells("A1:H1")
    t = ws["A1"]
    t.value     = f"{division_name} — 2026 Pipeline Waterfall"
    t.font      = Font(bold=True, size=14, color="E8521A")
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 6

    # Header for the visible data table (cols A-B)
    for col, hdr in enumerate(["Category", "Value"], 1):
        c = ws.cell(row=3, column=col, value=hdr)
        c.font      = Font(bold=True, color="FFFFFF", size=10)
        c.fill      = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 18

    # Data rows for the visible table
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

    # Summary block
    sr = data_end + 3
    ws.cell(sr, 1, "Summary").font = Font(bold=True, size=11, color="374151")
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

    # ── Chart data block (cols J & K) ────────────────────────────────────────
    # openpyxl must reference actual cells for the chart; we write labels and
    # values here. The numRef→strRef patch below fixes the axis display.
    chart_col_lbl = 10   # J
    chart_col_val = 11   # K
    chart_row_start = 3
    for i, (lbl, val) in enumerate(zip(labels, values)):
        ws.cell(row=chart_row_start + i, column=chart_col_lbl, value=lbl)
        ws.cell(row=chart_row_start + i, column=chart_col_val, value=val)
    chart_row_end = chart_row_start + len(labels) - 1

    # ── Build the chart ───────────────────────────────────────────────────────
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

    chart.add_data(
        Reference(ws, min_col=chart_col_val,
                  min_row=chart_row_start, max_row=chart_row_end),
        titles_from_data=False
    )
    # set_categories writes <numRef> — we'll patch it to <strRef> after saving
    chart.set_categories(
        Reference(ws, min_col=chart_col_lbl,
                  min_row=chart_row_start, max_row=chart_row_end)
    )

    # Per-bar colours
    ser = chart.series[0]
    ser.graphicalProperties.solidFill      = "E8521A"
    ser.graphicalProperties.line.solidFill = "FFFFFF"
    for i, lbl in enumerate(labels):
        dp = DataPoint(idx=i)
        hc = COLOR_MAP.get(lbl, "9CA3AF")
        dp.graphicalProperties.solidFill      = hc
        dp.graphicalProperties.line.solidFill = hc
        ser.dPt.append(dp)

    # Value labels above each bar
    dl = DataLabelList()
    dl.showVal       = True
    dl.showCatName   = False
    dl.showSerName   = False
    dl.showPercent   = False
    dl.showLegendKey = False
    dl.position      = "outEnd"
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

    risk_fills = {"low":"D1FAE5","medium":"FEF3C7","high":"FEE2E2","upside":"EDE9FE","2027+":"E0F2FE"}
    for ri, s in enumerate(sites, 2):
        vals = [s.get("sipId",""), s.get("restNum",""), s.get("fz",""),
                s.get("address",""), s.get("city",""), s.get("state",""),
                s.get("status",""), s.get("fzOpenDate",""), s.get("plkOpenDate",""),
                s.get("riskLevel",""), s.get("lastComment","")]
        for col, val in enumerate(vals, 1):
            cell = ws2.cell(row=ri, column=col, value=val)
            cell.alignment = Alignment(vertical="center", wrap_text=(col == 11))
        ws2.row_dimensions[ri].height = 16
        rf = risk_fills.get(s.get("riskLevel","").strip().lower())
        if rf:
            ws2.cell(row=ri, column=10).fill = PatternFill("solid", fgColor=rf)

    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(hdrs))}1"

    # ── Save to buffer ────────────────────────────────────────────────────────
    pre_buf = io.BytesIO()
    wb.save(pre_buf)
    pre_buf.seek(0)

    # ── Patch chart XML: numRef → strRef for X-axis category labels ───────────
    # openpyxl always writes <c:numRef> for chart categories even when the cells
    # contain strings. Excel interprets this as numeric tick positions (1, 2, 3…)
    # so the stage names never appear. We rewrite the ZIP in-memory, replacing
    # the <c:cat><c:numRef>…</c:numRef></c:cat> block with a proper <c:strRef>
    # that embeds the label strings directly in a <c:strCache>.
    # openpyxl serialises chart XML with NO namespace prefix on its own elements,
    # so tags are <cat>, <numRef>, <f>, etc. — NOT <c:cat>, <c:numRef>, <c:f>.
    pt_tags = "".join(
        f'<pt idx="{i}"><v>{lbl}</v></pt>'
        for i, lbl in enumerate(labels)
    )
    str_cache = (
        f'<strCache>'
        f'<ptCount val="{len(labels)}"/>'
        f'{pt_tags}'
        f'</strCache>'
    )

    out_buf = io.BytesIO()
    with zipfile.ZipFile(pre_buf, "r") as zin, \
         zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/charts/chart1.xml":
                xml = data.decode("utf-8")
                # Replace <cat><numRef>…</numRef></cat>
                # with    <cat><strRef><f>…</f><strCache>…</strCache></strRef></cat>
                def _patch_cat(m):
                    inner = m.group(1)
                    f_match = re.search(r'<f>(.*?)</f>', inner, re.DOTALL)
                    f_tag = f_match.group(0) if f_match else ""
                    return f'<cat><strRef>{f_tag}{str_cache}</strRef></cat>'

                xml = re.sub(
                    r'<cat>\s*<numRef>(.*?)</numRef>\s*</cat>',
                    _patch_cat,
                    xml,
                    flags=re.DOTALL
                )
                data = xml.encode("utf-8")
            zout.writestr(item, data)

    out_buf.seek(0)
    return out_buf.read()


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
