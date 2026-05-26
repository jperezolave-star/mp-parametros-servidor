import os, json, tempfile, subprocess, base64
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse

PORT = int(os.environ.get("PORT", 3000))

def get_drive_service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    creds_json = json.loads(os.environ["GOOGLE_CREDENTIALS"])
    creds = service_account.Credentials.from_service_account_info(
        creds_json, scopes=["https://www.googleapis.com/auth/drive.file"])
    return build("drive", "v3", credentials=creds)

def subir_drive(service, file_path, file_name, folder_id, mime_type):
    from googleapiclient.http import MediaFileUpload
    file_metadata = {"name": file_name, "parents": [folder_id]}
    media = MediaFileUpload(file_path, mimetype=mime_type)
    f = service.files().create(body=file_metadata, media_body=media, fields="id,webViewLink").execute()
    return f

class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        print(format % args)

    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(json.dumps({
            "status": "ok",
            "app": "MP Ingeniería — Servidor de parámetros",
            "version": "1.0.0"
        }).encode())

    def do_POST(self):
        if urlparse(self.path).path != "/generar":
            self.send_response(404)
            self.end_headers()
            return

        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length)
        datos = json.loads(body)

        tmp = tempfile.mkdtemp()
        try:
            json_path = os.path.join(tmp, "datos.json")
            with open(json_path, "w") as f:
                json.dump(datos, f)

            mes  = (datos.get("mes") or "MES").upper().replace(" ", "_")
            anio = datos.get("anio") or "2026"
            slug = (datos.get("edificio") or "EDIFICIO")
            slug = "".join(c if c.isalnum() or c == "_" else "_" for c in slug.replace(" ", "_"))[:20]

            # Generar Excel
            excel_out = os.path.join(tmp, f"PARAMETROS_{slug}_{mes}_{anio}.xlsx")
            subprocess.run(["python3", "/app/generar_excel.py", json_path, excel_out],
                           check=True, timeout=60)

            # Generar gráficos
            graf_dir = os.path.join(tmp, "graficos")
            os.makedirs(graf_dir, exist_ok=True)
            graf_result = subprocess.run(
                ["python3", "/app/generar_graficos.py", graf_dir],
                input=open(json_path).read().encode(),
                capture_output=True, timeout=60)
            try:
                graf_paths = json.loads(graf_result.stdout.decode().strip())
            except:
                graf_paths = {}

            datos["_graf_paths"] = graf_paths
            with open(json_path, "w") as f:
                json.dump(datos, f)

            # Generar informe Word
            word_out = os.path.join(tmp, f"INFORME_{slug}_{mes}_{anio}.docx")
            subprocess.run(["node", "/app/generar_informe.js", json_path, word_out],
                           check=True, timeout=120)

            # Subir a Drive
            service = get_drive_service()
            folder_excel = os.environ["DRIVE_FOLDER_EXCEL"]
            folder_word  = os.environ["DRIVE_FOLDER_WORD"]

            excel_file = subir_drive(service, excel_out,
                os.path.basename(excel_out), folder_excel,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            word_file = subir_drive(service, word_out,
                os.path.basename(word_out), folder_word,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

            resp = {
                "ok": True,
                "excel_url": excel_file.get("webViewLink"),
                "word_url":  word_file.get("webViewLink"),
            }
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps(resp).encode())

        except Exception as e:
            import traceback
            print("ERROR:", traceback.format_exc())
            self.send_response(500)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"ok": False, "error": str(e)}).encode())
        finally:
            import shutil
            shutil.rmtree(tmp, ignore_errors=True)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

print(f"✓ Servidor MP Ingeniería corriendo en puerto {PORT}")
HTTPServer(("0.0.0.0", PORT), Handler).serve_forever()
