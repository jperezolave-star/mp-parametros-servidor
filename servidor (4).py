import os, json, tempfile, subprocess, shutil, sys
from flask import Flask, request, jsonify, make_response

app = Flask(__name__)
PORT = int(os.environ.get("PORT", 8080))

# Instalar Node.js al arrancar si no está disponible
def instalar_node():
    try:
        subprocess.run(["node", "--version"], check=True, capture_output=True)
        print("✓ Node.js ya instalado")
    except FileNotFoundError:
        print("Instalando Node.js...")
        subprocess.run(
            "curl -fsSL https://deb.nodesource.com/setup_20.x | bash - && apt-get install -y nodejs",
            shell=True, check=True)
        subprocess.run(["npm", "install", "-g", "docx"], check=True)
        print("✓ Node.js instalado")

instalar_node()

@app.after_request
def add_cors(response):
    response.headers["Access-Control-Allow-Origin"]  = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type, Accept"
    return response

def get_drive_service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    creds_json = json.loads(os.environ["GOOGLE_CREDENTIALS"])
    creds = service_account.Credentials.from_service_account_info(
        creds_json, scopes=["https://www.googleapis.com/auth/drive.file"])
    return build("drive", "v3", credentials=creds)

def subir_drive(service, file_path, file_name, folder_id, mime_type):
    from googleapiclient.http import MediaFileUpload
    meta  = {"name": file_name, "parents": [folder_id]}
    media = MediaFileUpload(file_path, mimetype=mime_type)
    f = service.files().create(body=meta, media_body=media, fields="id,webViewLink").execute()
    return f

@app.route("/", methods=["GET"])
def health():
    return jsonify({"status": "ok", "app": "MP Ingeniería — Servidor de parámetros", "version": "1.0.0"})

@app.route("/generar", methods=["OPTIONS"])
def generar_options():
    return make_response("", 200)

@app.route("/generar", methods=["POST"])
def generar():
    datos = request.get_json()
    tmp   = tempfile.mkdtemp()
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

        return jsonify({
            "ok":        True,
            "excel_url": excel_file.get("webViewLink"),
            "word_url":  word_file.get("webViewLink"),
        })

    except Exception as e:
        import traceback
        print("ERROR:", traceback.format_exc())
        return jsonify({"ok": False, "error": str(e)}), 500
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

if __name__ == "__main__":
    print(f"✓ Servidor MP Ingeniería corriendo en puerto {PORT}")
    app.run(host="0.0.0.0", port=PORT)
