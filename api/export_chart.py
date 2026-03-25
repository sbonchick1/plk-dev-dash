import json
import io
import zipfile
import re
from http.server import BaseHTTPRequestHandler
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

COLOR_MAP = {
    "Prospect":"E8521A","SA":"E8521A","PC":"E8521A",
    "Permitting":"E8521A","UC":"E8521A","Open":"E8521A",
    "FY BU":"7B2D8B","Upside":"C1272D","Budget":"00A99D",
}
HEADER_FILL = PatternFill("solid", fgColor="374151")
STRIPE_FILL = PatternFill("solid", fgColor="F3F4F6")
WHITE_FILL  = PatternFill("solid", fgColor="FFFFFF")
STATUSES    = ["Prospect","SA","PC","Permitting","UC","Open"]


def patch_chart_xml(xml, str_cache):
    # Fix 1: Remove invalid <tx><v>string</v></tx> series title tags
    # (openpyxl serializes Series(title="...") as <tx><v>...</v></tx> which is
    #  invalid OOXML — Excel drops the entire chart on recovery when it sees this)
    xml = re.sub(r'<tx><v>[^<]*</v></tx>', '', xml)

    # Fix 2: numRef -> strRef for X-axis category labels
    def _replace_cat(m):
        inner = m.group(1)
        f_match = re.search(r'<f>(.*?)</f>', inner, re.DOTALL)
        f_tag = f_match.group(0) if f_match else ""
        return '<cat><strRef>' + f_tag + str_cache + '</strRef></cat>'
    xml = re.sub(r'<cat>\s*<numRef>(.*?)</numRef>\s*</cat>', _replace_cat, xml, flags=re.DOTALL)

    # Fix 3: Inject required namespaces onto chartSpace root
    old_root = 'xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart">'
    new_root = (
        'xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    )
    xml = xml.replace('<chartSpace ' + old_root, '<chartSpace ' + new_root, 1)
    xml = re.sub(r' xmlns:a="http://schemas\.openxmlformats\.org/drawingml/2006/main"', '', xml)
    if 'xmlns:a' not in xml[:400]:
        xml = xml.replace(
            'xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
            'xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart"'
            ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
            1)
    return xml


def rezip_with_correct_order(raw_xlsx_bytes, chart_xml_patch_fn):
    with zipfile.ZipFile(io.BytesIO(raw_xlsx_bytes), 'r') as zin:
        all_files = {}
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'xl/charts/chart1.xml':
                data = chart_xml_patch_fn(data.decode('utf-8')).encode('utf-8')
            all_files[item.filename] = data
    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in ['[Content_Types].xml', '_rels/.rels']:
            if name in all_files:
                zout.writestr(name, all_files[name])
        for name, data in all_files.items():
            if name not in ('[Content_Types].xml', '_rels/.rels'):
                zout.writestr(name, data)
    out_buf.seek(0)
    return out_buf.read()


def build_xlsx(payload):
    division_name = payload.get("divisionName", "Division")
    labels  = payload.get("labels", [])
    values  = payload.get("displayValues", [])
    budget  = payload.get("budget", 0)
    fy_bu   = payload.get("fyBU", 0)
    upside  = payload.get("upsideCount", 0)
    gap     = payload.get("gap", 0)
    sites   = payload.get("sites", [])

    bar_labels = STATUSES + ["FY BU", "Upside", "Budget"]
    n = len(bar_labels)

    waterfall_vals = []
    sip_vals = []
    cumulative = 0
    for lbl in STATUSES:
        idx = labels.index(lbl) if lbl in labels else -1
        v = values[idx] if idx >= 0 else 0
        waterfall_vals.append(cumulative)
        sip_vals.append(v)
        cumulative += v
    waterfall_vals.append(0);     sip_vals.append(fy_bu)
    waterfall_vals.append(fy_bu); sip_vals.append(upside)
    waterfall_vals.append(0);     sip_vals.append(budget)

    wb = Workbook()
    ws = wb.active
    ws.title = "Waterfall Chart"
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:C1")
    t = ws["A1"]
    t.value = division_name + " \u2014 2026 Pipeline Waterfall"
    t.font = Font(bold=True, size=14, color="E8521A")
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 6

    for col, hdr in enumerate(["Stage", "Waterfall", "# SIPs"], 1):
        c = ws.cell(row=3, column=col, value=hdr)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 18

    table_row = 4
    for i, lbl in enumerate(bar_labels):
        fill = STRIPE_FILL if i % 2 == 0 else WHITE_FILL
        hc = COLOR_MAP.get(lbl, "6B7280")
        r = table_row + i
        c1 = ws.cell(row=r, column=1, value=lbl)
        c1.font = Font(bold=(lbl in ("FY BU","Budget","Upside")), color=hc)
        c1.fill = fill
        c1.alignment = Alignment(horizontal="left", vertical="center")
        c2 = ws.cell(row=r, column=2, value=waterfall_vals[i])
        c2.font = Font(color="9CA3AF"); c2.fill = fill
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c2.number_format = "0"
        c3 = ws.cell(row=r, column=3, value=sip_vals[i])
        c3.font = Font(bold=True, color=hc); c3.fill = fill
        c3.alignment = Alignment(horizontal="center", vertical="center")
        c3.number_format = "0"
        ws.row_dimensions[r].height = 17

    table_end = table_row + n - 1

    sr = table_end + 2
    gc = "059669" if gap >= 0 else "DC2626"
    def stat_row(row, label, value, fmt="0", color="374151"):
        ws.cell(row=row, column=1, value=label).font = Font(color="6B7280")
        c = ws.cell(row=row, column=2, value=value)
        c.font = Font(bold=True, color=color)
        c.number_format = fmt
    stat_row(sr,   "FY BU",                 fy_bu)
    stat_row(sr+1, "Budget",                budget)
    stat_row(sr+2, "Gap (FY BU vs Budget)", gap,                         "+0;-0;0", gc)
    stat_row(sr+3, "Gap %",                 gap/budget if budget else 0, "0%",      gc)
    stat_row(sr+4, "FY BU + Upside",        fy_bu + upside)

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 10

    chart = BarChart()
    chart.type     = "col"
    chart.grouping = "stacked"
    chart.overlap  = 100
    chart.title    = division_name + " \u2014 2026 Pipeline"
    chart.width    = 26
    chart.height   = 16
    chart.legend   = None
    chart.y_axis.majorGridlines = None
    chart.y_axis.delete         = True
    chart.x_axis.delete         = False
    chart.x_axis.tickLblPos     = "low"
    chart.x_axis.numFmt         = "General"
    chart.x_axis.axPos          = "b"

    # Note: do NOT set title= on Series — openpyxl serializes it as
    # invalid <tx><v>string</v></tx> which causes Excel to drop the chart
    base_ref = Reference(ws, min_col=2, min_row=table_row, max_row=table_end)
    base_ser = Series(base_ref)
    base_ser.graphicalProperties.solidFill      = "FFFFFF"
    base_ser.graphicalProperties.line.solidFill = "FFFFFF"
    base_ser.graphicalProperties.line.width     = 0
    chart.append(base_ser)

    bar_ref = Reference(ws, min_col=3, min_row=table_row, max_row=table_end)
    bar_ser = Series(bar_ref)
    bar_ser.graphicalProperties.solidFill      = "E8521A"
    bar_ser.graphicalProperties.line.solidFill = "FFFFFF"
    bar_ser.graphicalProperties.line.width     = 6350
    for i, lbl in enumerate(bar_labels):
        dp = DataPoint(idx=i)
        hc = COLOR_MAP.get(lbl, "9CA3AF")
        dp.graphicalProperties.solidFill      = hc
        dp.graphicalProperties.line.solidFill = hc
        bar_ser.dPt.append(dp)
    dl = DataLabelList()
    dl.showVal=True; dl.showCatName=False; dl.showSerName=False
    dl.showPercent=False; dl.showLegendKey=False; dl.position="outEnd"
    bar_ser.dLbls = dl
    chart.append(bar_ser)

    cats = Reference(ws, min_col=1, min_row=table_row, max_row=table_end)
    chart.set_categories(cats)
    ws.add_chart(chart, "E2")

    ws2 = wb.create_sheet("Site Detail")
    ws2.sheet_view.showGridLines = False
    hdrs   = ["SIP ID","Rest No","FZ","Address","City","ST","Status",
              "FZ Proj Open Date","PLK Proj Open Date","Risk Level","Last Comments"]
    widths = [12,10,20,24,16,6,12,18,18,12,44]
    for col, (h, w) in enumerate(zip(hdrs, widths), 1):
        c = ws2.cell(row=1, column=col, value=h)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws2.column_dimensions[get_column_letter(col)].width = w
    ws2.row_dimensions[1].height = 18
    risk_fills = {"low":"D1FAE5","medium":"FEF3C7","high":"FEE2E2","upside":"EDE9FE","2027+":"E0F2FE"}
    for ri, s in enumerate(sites, 2):
        row_vals = [s.get("sipId",""),s.get("restNum",""),s.get("fz",""),
                    s.get("address",""),s.get("city",""),s.get("state",""),
                    s.get("status",""),s.get("fzOpenDate",""),s.get("plkOpenDate",""),
                    s.get("riskLevel",""),s.get("lastComment","")]
        for col, val in enumerate(row_vals, 1):
            cell = ws2.cell(row=ri, column=col, value=val)
            cell.alignment = Alignment(vertical="center", wrap_text=(col==11))
        ws2.row_dimensions[ri].height = 16
        rf = risk_fills.get(s.get("riskLevel","").strip().lower())
        if rf:
            ws2.cell(row=ri, column=10).fill = PatternFill("solid", fgColor=rf)
    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = "A1:" + get_column_letter(len(hdrs)) + "1"

    pre_buf = io.BytesIO()
    wb.save(pre_buf)
    pre_buf.seek(0)

    pt_tags = "".join('<pt idx="'+str(i)+'"><v>'+lbl+'</v></pt>' for i, lbl in enumerate(bar_labels))
    str_cache = '<strCache><ptCount val="'+str(n)+'"/>'+pt_tags+'</strCache>'

    def apply_patches(xml_str):
        return patch_chart_xml(xml_str, str_cache)

    return rezip_with_correct_order(pre_buf.read(), apply_patches)


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
            self._respond_error(400, "Bad request: " + str(e))
            return
        try:
            xlsx_bytes = build_xlsx(payload)
        except Exception as e:
            self._respond_error(500, str(e))
            return

        division_name = payload.get("divisionName", "Division")
        safe     = "".join(c if c.isalnum() or c in " _-" else "_" for c in division_name)
        filename = "Waterfall_" + safe + ".xlsx"

        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", 'attachment; filename="' + filename + '"')
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
