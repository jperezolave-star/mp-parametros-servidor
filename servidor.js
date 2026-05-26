const express = require("express");
const cors    = require("cors");
const fs      = require("fs");
const path    = require("path");
const { execSync } = require("child_process");
const { google } = require("googleapis");

const app  = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: "50mb" }));

// ── Google Drive setup ────────────────────────────────────────────────────
function getDriveClient() {
  const creds = JSON.parse(process.env.GOOGLE_CREDENTIALS);
  const auth  = new google.auth.GoogleAuth({
    credentials: creds,
    scopes: ["https://www.googleapis.com/auth/drive.file"],
  });
  return google.drive({ version: "v3", auth });
}

async function subirDrive(drive, filePath, fileName, folderId, mimeType) {
  const res = await drive.files.create({
    requestBody: {
      name:    fileName,
      parents: [folderId],
    },
    media: {
      mimeType,
      body: fs.createReadStream(filePath),
    },
    fields: "id, webViewLink",
  });
  return res.data;
}

// ── Ruta principal: generar Excel + Word y subir a Drive ──────────────────
app.post("/generar", async (req, res) => {
  const datos = req.body;
  const ts    = Date.now();
  const tmpDir = `/tmp/mp_${ts}`;
  fs.mkdirSync(tmpDir, { recursive: true });

  try {
    // 1. Guardar datos JSON
    const jsonPath = path.join(tmpDir, "datos.json");
    fs.writeFileSync(jsonPath, JSON.stringify(datos, null, 2));

    const mes   = (datos.mes  || "MES").toUpperCase().replace(/\s+/g, "_");
    const anio  = datos.anio  || "2026";
    const slug  = (datos.edificio || "EDIFICIO").replace(/\s+/g,"_").replace(/[^A-Za-z0-9_]/g,"").slice(0,20);

    // 2. Generar Excel
    console.log("Generando Excel...");
    const excelOut = path.join(tmpDir, `PARAMETROS_${slug}_${mes}_${anio}.xlsx`);
    execSync(`python3 /app/generar_excel.py ${jsonPath} ${excelOut}`, { timeout: 60000 });

    // 3. Generar gráficos
    console.log("Generando gráficos...");
    const grafDir = path.join(tmpDir, "graficos");
    fs.mkdirSync(grafDir, { recursive: true });
    const grafResult = execSync(
      `python3 /app/generar_graficos.py ${grafDir} < ${jsonPath}`,
      { timeout: 60000, maxBuffer: 10 * 1024 * 1024 }
    );
    const grafPaths = JSON.parse(grafResult.toString().trim());

    // Agregar paths al JSON para el informe
    const datosConGraf = { ...datos, _graf_paths: grafPaths };
    fs.writeFileSync(jsonPath, JSON.stringify(datosConGraf, null, 2));

    // 4. Generar informe Word
    console.log("Generando informe Word...");
    const wordOut = path.join(tmpDir, `INFORME_${slug}_${mes}_${anio}.docx`);
    execSync(`node /app/generar_informe.js ${jsonPath} ${wordOut}`, { timeout: 120000 });

    // 5. Subir a Google Drive
    console.log("Subiendo a Drive...");
    const drive         = getDriveClient();
    const folderExcel   = process.env.DRIVE_FOLDER_EXCEL;
    const folderWord    = process.env.DRIVE_FOLDER_WORD;

    const [excelFile, wordFile] = await Promise.all([
      subirDrive(drive, excelOut, path.basename(excelOut), folderExcel,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
      subirDrive(drive, wordOut,  path.basename(wordOut),  folderWord,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
    ]);

    console.log("✓ Archivos subidos a Drive");

    // 6. Limpiar temporales
    fs.rmSync(tmpDir, { recursive: true, force: true });

    res.json({
      ok:        true,
      excel_url: excelFile.webViewLink,
      word_url:  wordFile.webViewLink,
      excel_id:  excelFile.id,
      word_id:   wordFile.id,
    });

  } catch (err) {
    console.error("Error al generar:", err.message);
    fs.rmSync(tmpDir, { recursive: true, force: true });
    res.status(500).json({ ok: false, error: err.message });
  }
});

// ── Health check ──────────────────────────────────────────────────────────
app.get("/", (req, res) => {
  res.json({
    status:    "ok",
    app:       "MP Ingeniería — Servidor de parámetros",
    version:   "1.0.0",
    timestamp: new Date().toISOString(),
  });
});

app.listen(PORT, () => {
  console.log(`✓ Servidor MP Ingeniería corriendo en puerto ${PORT}`);
});
