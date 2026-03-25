from http.server import BaseHTTPRequestHandler
import json, sys

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        result = {"python": sys.version, "status": "ok"}
        
        try:
            import openpyxl
            result["openpyxl"] = openpyxl.__version__
        except Exception as e:
            result["openpyxl_error"] = str(e)

        try:
            from openpyxl.chart.label import DataLabelList
            result["DataLabelList"] = "ok"
        except Exception as e:
            result["DataLabelList_error"] = str(e)

        try:
            from openpyxl.chart.series import DataPoint
            result["DataPoint"] = "ok"
        except Exception as e:
            result["DataPoint_error"] = str(e)

        body = json.dumps(result, indent=2).encode()
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):
        pass
