import os, json, tempfile, subprocess, shutil
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse

PORT = int(os.environ.get("PORT", 8080))

CORS_HEADERS = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
}

def get_drive_service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    creds_json = json.loads(os.environ["GOOGLE_CREDENTIALS"])
    creds = service_account.Credentials.from_service_account_info(
        creds_json, scopes=["https://www.googleapis.com/auth/drive.file"])
    return build("drive", "v3", credentials=creds)

def subir_drive(service, file_path, file_name, folder_id, mime_type):
    from googleapiclient.http import MediaFileUpload
    meta = {"name": file_name, "parents": [folder_id]}
    media = MediaFileUpload(file_path, mimetype=mime_type)
    f = service.files().create(body=meta, media_body=media, fields="id,webViewLink").execute()
    return f

def send_json(handler, code, data):
    body = json.dumps(data, ensure_ascii=False).encode()
    handler.send_response(code)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    for k, v in CORS_HEADERS.items():
        handler.send_header(k, v)
    handler.end_headers()
    handler.wfile.write(body)

class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        print(fmt % args)

    def do_OPTIONS(self):
        self.send_response(200)
        for k, v in CORS_HEADERS.items():
            self.send_header(k, v)
        self.end_headers()

    def do_GET(self):
        send_json(self, 200, {
            "status": "ok",
            "app": "MP Ingeniería — Servidor de parámetros",
            "version": "1.0.0"
        })

    def do_POST(self):
        if urlparse(self.path).path != "/generar":
            send_json(self, 404, {"error": "Not found"})
            return

        length = int(self.headers.get("Content-Length", 0))
        body   = self.rfile.read(length)
        datos  = json.loads(body)

        tmp = tempfile.mkdtemp()
        try:
            json_path = os.path.join(tmp, "datos.json")
            with open(json_path, "w") as f:
                json.dump(datos, f, ensure_ascii=False)

            mes  = (datos.get("mes") or "MES").upper().replace(" ", "_")
            anio = datos.get("anio") or "2026"
            slug = "".join(c if c.isalnum() or c=="_" else "_"
                           for c in (datos.get("edificio") or "EDIFICIO").replace(" ","_"))[:20]

            # Generar Excel
            excel_out = os.path.join(tmp, f"PARAMETROS_{slug}_{mes}_{anio}.xlsx")
            subprocess.run(["python3", "/app/generar_excel.py", json_path, excel_out],
                           check=True, timeout=60)

            # Generar gráficos
            graf_dir = os.path.join(tmp, "graficos")
            os.makedirs(graf_dir, exist_ok=True)
            gr = subprocess.run(
                ["python3", "/app/generar_graficos.py", graf_dir],
                input=open(json_path, "rb").read(),
                capture_output=True, timeout=60)
            try:
                graf_paths = json.loads(gr.stdout.decode().strip())
            except:
                graf_paths = {}

            datos["_graf_paths"] = graf_paths
            with open(json_path, "w") as f:
                json.dump(datos, f, ensure_ascii=False)

            # Generar informe Word
            word_out = os.path.join(tmp, f"INFORME_{slug}_{mes}_{anio}.docx")
            subprocess.run(["node", "/app/generar_informe.js", json_path, word_out],
                           check=True, timeout=120)

            # Subir a Drive
            service      = get_drive_service()
            folder_excel = os.environ["DRIVE_FOLDER_EXCEL"]
            folder_word  = os.environ["DRIVE_FOLDER_WORD"]

            excel_file = subir_drive(service, excel_out, os.path.basename(excel_out),
                folder_excel,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            word_file  = subir_drive(service, word_out, os.path.basename(word_out),
                folder_word,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

            send_json(self, 200, {
                "ok":        True,
                "excel_url": excel_file.get("webViewLink"),
                "word_url":  word_file.get("webViewLink"),
            })

        except Exception as e:
            import traceback
            print("ERROR:", traceback.format_exc())
            send_json(self, 500, {"ok": False, "error": str(e)})
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

print(f"✓ Servidor MP Ingeniería corriendo en puerto {PORT}")
HTTPServer(("0.0.0.0", PORT), Handler).serve_forever()
