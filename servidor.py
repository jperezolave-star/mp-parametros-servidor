import os, json, tempfile, subprocess, shutil
import urllib.request, urllib.error
from flask import Flask, request, jsonify, make_response

app = Flask(__name__)
PORT = int(os.environ.get("PORT", 8080))

SUPABASE_URL = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "")

# Valores por defecto si Supabase no responde
DEFAULT_CONFIG = {
    "tecnicos": [
        "Acevedo Donoso, Mauricio Eugenio",
        "Arango Bravo, Jaime Hernán",
        "Bahamondes Palome, Claudio Mario",
        "Bustillos Pradenas, Benjamín",
        "De la Cruz Carreño Miranda, Alfonso Albino",
        "González Conopoima, Efraín Manuel",
        "Hermosilla Vega, Ricardo Manuel",
        "Rivera Navarro, Ismael Luciano",
        "Sánchez Soto, César Fernando",
        "Torres Yáñez, Kevin Daniel",
        "Uribe López, Luis Mario"
    ],
    "jefes": [
        "Jorge Rivera Navarro",
        "Jose Perez Olave",
        "Nelson Mena Malet"
    ]
}

def sb_get(clave):
    """Lee un valor de la tabla config en Supabase."""
    try:
        url = f"{SUPABASE_URL}/rest/v1/config?clave=eq.{clave}&select=valor"
        req = urllib.request.Request(url, headers={
            "apikey": SUPABASE_KEY,
            "Authorization": f"Bearer {SUPABASE_KEY}"
        })
        with urllib.request.urlopen(req, timeout=5) as r:
            data = json.loads(r.read())
            if data:
                return data[0]["valor"]
    except Exception as e:
        print(f"Supabase GET error ({clave}):", e)
    return None

def sb_set(clave, valor):
    """Guarda un valor en la tabla config en Supabase (upsert)."""
    try:
        url = f"{SUPABASE_URL}/rest/v1/config"
        payload = json.dumps({"clave": clave, "valor": valor}).encode()
        req = urllib.request.Request(url, data=payload, method="POST", headers={
            "apikey": SUPABASE_KEY,
            "Authorization": f"Bearer {SUPABASE_KEY}",
            "Content-Type": "application/json",
            "Prefer": "resolution=merge-duplicates"
        })
        urllib.request.urlopen(req, timeout=5)
        return True
    except Exception as e:
        print(f"Supabase SET error ({clave}):", e)
    return False

@app.after_request
def add_cors(response):
    response.headers["Access-Control-Allow-Origin"]  = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS, PUT, DELETE"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type, Accept"
    return response

@app.errorhandler(Exception)
def handle_exception(e):
    import traceback
    print("ERROR:", traceback.format_exc())
    resp = jsonify({"ok": False, "error": str(e)})
    resp.status_code = 500
    resp.headers["Access-Control-Allow-Origin"] = "*"
    return resp

def get_drive_service():
    import google.oauth2.credentials
    import google.auth.transport.requests
    from googleapiclient.discovery import build
    creds = google.oauth2.credentials.Credentials(
        token=None,
        refresh_token=os.environ["GOOGLE_REFRESH_TOKEN"],
        client_id=os.environ["GOOGLE_CLIENT_ID"],
        client_secret=os.environ["GOOGLE_CLIENT_SECRET"],
        token_uri="https://oauth2.googleapis.com/token",
        scopes=["https://www.googleapis.com/auth/drive.file"],
        enable_reauth_refresh=False
    )
    request_session = google.auth.transport.requests.Request()
    creds.refresh(request_session)
    return build("drive", "v3", credentials=creds)

def subir_drive(service, file_path, file_name, folder_id, mime_type):
    from googleapiclient.http import MediaFileUpload
    meta  = {"name": file_name, "parents": [folder_id]}
    media = MediaFileUpload(file_path, mimetype=mime_type)
    f = service.files().create(body=meta, media_body=media, fields="id,webViewLink").execute()
    return f

# ── Endpoints de configuración global ──────────────────────────────────────

@app.route("/config", methods=["OPTIONS"])
def config_options():
    return make_response("", 200)

@app.route("/config", methods=["GET"])
def get_config():
    tecnicos = sb_get("tecnicos") or DEFAULT_CONFIG["tecnicos"]
    jefes    = sb_get("jefes")    or DEFAULT_CONFIG["jefes"]
    return jsonify({"ok": True, "tecnicos": tecnicos, "jefes": jefes})

@app.route("/config", methods=["POST"])
def save_config():
    datos = request.get_json()
    if "tecnicos" in datos:
        sb_set("tecnicos", datos["tecnicos"])
    if "jefes" in datos:
        sb_set("jefes", datos["jefes"])
    print(f"✓ Config guardada en Supabase")
    return jsonify({"ok": True})

# ── Endpoint principal ──────────────────────────────────────────────────────

@app.route("/", methods=["GET"])
def health():
    return jsonify({"status":"ok","app":"MP Ingeniería — Servidor de parámetros","version":"2.1.0"})

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

        excel_out = os.path.join(tmp, f"PARAMETROS_{slug}_{mes}_{anio}.xlsx")
        subprocess.run(["python3", "/app/generar_excel.py", json_path, excel_out],
                       check=True, timeout=60)
        print(f"✓ Excel generado")

        graf_dir = os.path.join(tmp, "graficos")
        os.makedirs(graf_dir, exist_ok=True)
        gr = subprocess.run(["python3", "/app/generar_graficos.py", graf_dir],
            input=open(json_path,"rb").read(), capture_output=True, timeout=60)
        try: graf_paths = json.loads(gr.stdout.decode().strip())
        except: graf_paths = {}
        datos["_graf_paths"] = graf_paths
        with open(json_path, "w") as f:
            json.dump(datos, f, ensure_ascii=False)
        print(f"✓ Gráficos: {list(graf_paths.keys())}")

        word_out = os.path.join(tmp, f"INFORME_{slug}_{mes}_{anio}.docx")
        subprocess.run(["python3", "/app/generar_informe.py", json_path, word_out],
                       check=True, timeout=120)
        print(f"✓ Word generado")

        service      = get_drive_service()
        folder_excel = os.environ["DRIVE_FOLDER_EXCEL"]
        folder_word  = os.environ["DRIVE_FOLDER_WORD"]

        excel_file = subir_drive(service, excel_out, os.path.basename(excel_out),
            folder_excel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        word_file  = subir_drive(service, word_out, os.path.basename(word_out),
            folder_word, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        print(f"✓ Archivos subidos a Drive")

        return jsonify({"ok":True,
                        "excel_url":excel_file.get("webViewLink"),
                        "word_url": word_file.get("webViewLink")})

    except Exception as e:
        import traceback
        print("ERROR:", traceback.format_exc())
        return jsonify({"ok":False,"error":str(e)}), 500
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

if __name__ == "__main__":
    print(f"✓ Servidor MP Ingeniería v2.1 corriendo en puerto {PORT}")
    app.run(host="0.0.0.0", port=PORT)
