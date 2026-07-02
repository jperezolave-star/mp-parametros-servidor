import sys, json, os, base64
from io import BytesIO
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image

# ── Rutas ──────────────────────────────────────────────────────────────────
LOGO_PATH = os.environ.get("LOGO_PATH", "/app/logo_transparente.png")

# ── Colores ────────────────────────────────────────────────────────────────
AZ   = RGBColor(0x1F, 0x4E, 0x79)
AZ2  = RGBColor(0x2E, 0x75, 0xB6)
GRIS = RGBColor(0xF2, 0xF2, 0xF2)
VDE  = RGBColor(0xE2, 0xEF, 0xDA)
BLCO = RGBColor(0xFF, 0xFF, 0xFF)
RJO  = RGBColor(0xC0, 0x00, 0x00)
AMR  = RGBColor(0x7F, 0x60, 0x00)

def rgb_hex(r,g,b): return f"{r:02X}{g:02X}{b:02X}"

AZ_HEX  = "1F4E79"
AZ2_HEX = "2E75B6"
GR_HEX  = "F2F2F2"
VD_HEX  = "E2EFDA"

# ── Helpers documento ──────────────────────────────────────────────────────
def set_cell_bg(cell, hex_color):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)

def set_cell_borders(cell, color="CCCCCC"):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top","left","bottom","right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"),  "single")
        el.set(qn("w:sz"),   "4")
        el.set(qn("w:space"),"0")
        el.set(qn("w:color"),color)
        tcBorders.append(el)
    tcPr.append(tcBorders)

def set_no_borders(cell):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top","left","bottom","right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"),  "none")
        el.set(qn("w:sz"),   "0")
        el.set(qn("w:space"),"0")
        el.set(qn("w:color"),"FFFFFF")
        tcBorders.append(el)
    tcPr.append(tcBorders)

def set_cell_margins(cell, top=60, bottom=60, left=100, right=100):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in [("top",top),("bottom",bottom),("left",left),("right",right)]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"),    str(val))
        el.set(qn("w:type"), "dxa")
        tcMar.append(el)
    tcPr.append(tcMar)

def cell_run(cell, text, bold=False, size=9, color=None, italic=False, align="center"):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if align=="center" else (WD_ALIGN_PARAGRAPH.LEFT if align=="left" else WD_ALIGN_PARAGRAPH.RIGHT)
    run = p.add_run(str(text) if text is not None else "—")
    run.bold  = bold
    run.font.size = Pt(size)
    run.font.name = "Arial"
    if color:
        run.font.color.rgb = color

def hdr_cell(cell, text, bg_hex=AZ2_HEX):
    set_cell_bg(cell, bg_hex)
    set_cell_borders(cell, "FFFFFF")
    cell_run(cell, text, bold=True, size=8, color=BLCO)

def val_cell(cell, text, bg_hex="FFFFFF", bold=False, size=8, color=None, align="center"):
    set_cell_bg(cell, bg_hex)
    set_cell_borders(cell)
    cell_run(cell, text if text not in (None,"") else "—", bold=bold, size=size,
             color=color or RGBColor(0,0,0), align=align)

def add_heading(doc, text, level=1, before=12, after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after  = Pt(after)
    run = p.add_run(text)
    run.bold = True
    run.font.name = "Arial"
    run.font.size = Pt(16 if level==1 else 12)
    run.font.color.rgb = AZ if level==1 else AZ2

def add_para(doc, text, size=9, before=0, after=4, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after  = Pt(after)
    run = p.add_run(text)
    run.font.name = "Arial"
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color

def set_col_widths(table, widths_cm):
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            if i < len(widths_cm):
                cell.width = Cm(widths_cm[i])

# ── Clasificación de alertas ───────────────────────────────────────────────
def calc_desb(r, s, t):
    try:
        vals = [float(v) for v in [r,s,t] if v is not None and str(v).strip()]
        if len(vals) < 3: return None
        avg = sum(vals) / 3
        if avg == 0: return None
        return (max(abs(v - avg) for v in vals) / avg) * 100
    except:
        return None

def clasif(tipo, val):
    try: v = float(val)
    except: return None
    if tipo == "tri":
        if 368 <= v <= 391: return None
        if (361 <= v <= 367) or (392 <= v <= 399): return "warn"
        return "crit"
    if tipo == "mono":
        if 213 <= v <= 227: return None
        if (209 <= v <= 212) or (228 <= v <= 231): return "warn"
        return "crit"
    if tipo == "ntie":
        if 0 <= v <= 2: return None
        if 2 < v <= 5: return "warn"
        return "crit"
    if tipo == "temp":
        if v <= 40: return None
        if v <= 50: return "warn"
        return "crit"
    if tipo == "desb":
        if v <= 3: return None
        if v <= 5: return "warn"
        return "crit"
    return None

def alertas_tablero(tb):
    probs = []
    checks = [
        ("V R/S","tri",tb.get("v_rs")), ("V R/T","tri",tb.get("v_rt")), ("V S/T","tri",tb.get("v_st")),
        ("V R/N","mono",tb.get("v_rn")), ("V S/N","mono",tb.get("v_sn")), ("V T/N","mono",tb.get("v_tn")),
        ("V R/Tie","ntie",tb.get("v_rtie")), ("V S/Tie","ntie",tb.get("v_stie")),
        ("V T/Tie","ntie",tb.get("v_ttie")), ("V N/Tie","ntie",tb.get("v_ntie")),
        ("T° R","temp",tb.get("t_r")), ("T° S","temp",tb.get("t_s")),
        ("T° T","temp",tb.get("t_t")), ("T° N","temp",tb.get("t_n")), ("T° Tie","temp",tb.get("t_tie")),
    ]
    for label, tipo, val in checks:
        s = clasif(tipo, val)
        if s:
            probs.append({"p": label, "v": str(val), "s": s})
    desb = tb.get("desbalance") or calc_desb(tb.get("i_r"), tb.get("i_s"), tb.get("i_t"))
    if desb is not None:
        s = clasif("desb", desb)
        if s:
            probs.append({"p": "Desbalance I", "v": f"{float(desb):.1f}%", "s": s})
    return probs

# ── Sección resumen de alertas ─────────────────────────────────────────────
def seccion_alertas(doc, tableros):
    add_heading(doc, "Resumen de alertas", level=1)
    con_alertas = [(tb, alertas_tablero(tb)) for tb in tableros]
    con_alertas = [(tb, probs) for tb, probs in con_alertas if probs]

    if not con_alertas:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run("✓  Todos los parámetros se encuentran dentro de los rangos normales.")
        run.bold = True; run.font.size = Pt(9); run.font.name = "Arial"
        run.font.color.rgb = RGBColor(0x05,0x96,0x69)
        doc.add_paragraph()
        return

    criticos = sum(1 for _,probs in con_alertas if any(p["s"]=="crit" for p in probs))
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(f"Se detectaron alertas en {len(con_alertas)} tablero(s): ")
    run.font.size = Pt(9); run.font.name = "Arial"
    run2 = p.add_run(f"{criticos} CRÍTICO(S)")
    run2.bold = True; run2.font.size = Pt(9); run2.font.name = "Arial"; run2.font.color.rgb = RJO
    run3 = p.add_run(f"  |  {len(con_alertas)-criticos} ADVERTENCIA(S)")
    run3.bold = True; run3.font.size = Pt(9); run3.font.name = "Arial"; run3.font.color.rgb = AMR

    tbl = doc.add_table(rows=1, cols=7)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdrs = ["N°","Tablero","Piso","Parámetro","Valor","Estado","Hora"]
    ws   = [0.8,4.5,0.8,2.5,1.2,2.2,1.5]
    for i,h in enumerate(hdrs):
        hdr_cell(tbl.rows[0].cells[i], h, AZ_HEX)
    set_col_widths(tbl, ws)

    for tb, probs in con_alertas:
        hora = (tb.get("fecha_hora") or "—").split(" ")[-1]
        for idx, prob in enumerate(probs):
            row = tbl.add_row()
            is_crit = prob["s"] == "crit"
            rf  = "FFF0F0" if is_crit else "FFFBEB"
            color_txt = RJO if is_crit else AMR
            status = "⚠ CRÍTICO" if is_crit else "⚡ ADVERTENCIA"
            sf  = "FFE4E4" if is_crit else "FFF3CD"
            val_cell(row.cells[0], str(tb.get("n","")) if idx==0 else "", rf)
            val_cell(row.cells[1], tb.get("nombre","") if idx==0 else "", rf, align="left")
            val_cell(row.cells[2], str(tb.get("piso","")) if idx==0 else "", rf)
            val_cell(row.cells[3], prob["p"], rf, align="left")
            val_cell(row.cells[4], prob["v"], rf, bold=True)
            val_cell(row.cells[5], status, sf, bold=True, color=color_txt)
            val_cell(row.cells[6], hora if idx==0 else "", rf, size=7)
            set_col_widths(tbl, ws)
    doc.add_paragraph()

# ── Tabla de un tablero ────────────────────────────────────────────────────
def tabla_tablero(doc, tb):
    # Fila encabezado
    tbl = doc.add_table(rows=1, cols=18)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    hrow = tbl.rows[0]
    # Merge todas las celdas del encabezado
    cell_h = hrow.cells[0]
    for i in range(1, 18):
        cell_h = cell_h.merge(hrow.cells[i])
    set_cell_bg(cell_h, AZ_HEX)
    set_cell_borders(cell_h, AZ_HEX)
    cell_h.text = ""
    p = cell_h.paragraphs[0]
    r1 = p.add_run(f"N° {tb.get('n','')}  —  {tb.get('nombre','')}")
    r1.bold = True; r1.font.size = Pt(9); r1.font.name = "Arial"; r1.font.color.rgb = BLCO
    r2 = p.add_run(f"   Piso {tb.get('piso','')}  |  {tb.get('servicio','')}")
    r2.font.size = Pt(8); r2.font.name = "Arial"; r2.font.color.rgb = RGBColor(0xBD,0xD7,0xEE)
    if tb.get("fecha_hora"):
        r3 = p.add_run(f"  ·  {tb['fecha_hora']}")
        r3.font.size = Pt(7); r3.font.name = "Arial"; r3.font.color.rgb = RGBColor(0x9D,0xC3,0xE6)
        r3.italic = True

    # Fila sub-encabezados grupos
    row2 = tbl.add_row()
    c0 = row2.cells[0].merge(row2.cells[5])  # Tensiones 6 cols
    hdr_cell(c0, "TENSIONES (V)", AZ2_HEX)
    c1 = row2.cells[6].merge(row2.cells[10]) # Corrientes 5 cols
    hdr_cell(c1, "CORRIENTES (A)", "1F6B3E")
    c2 = row2.cells[11].merge(row2.cells[15])# Temps 5 cols
    hdr_cell(c2, "TEMPERATURAS (°C)", "7F6000")
    c3 = row2.cells[16].merge(row2.cells[17])# Estado 2 cols
    hdr_cell(c3, "ESTADO", "843C0C")

    # Fila etiquetas
    lbl3 = tbl.add_row()
    labels = ["R/S","R/T","S/T","R/N","S/N","T/N","Cte R","Cte S","Cte T","Cte N","Cte TIE","T° R","T° S","T° T","T° N","T° TIE","Piloto","Chapas"]
    fills  = [GR_HEX]*6 + ["E2EFDA"]*5 + ["FFF2CC"]*5 + ["FCE4D6","FCE4D6"]
    for i,(lbl,fill) in enumerate(zip(labels,fills)):
        hdr_cell(lbl3.cells[i], lbl, fill)
        lbl3.cells[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(0x1F,0x4E,0x79)

    # Fila valores
    vrow = tbl.add_row()
    def v(x): return str(x) if x is not None else "—"
    vals = [
        v(tb.get("v_rs")),v(tb.get("v_rt")),v(tb.get("v_st")),
        v(tb.get("v_rn")),v(tb.get("v_sn")),v(tb.get("v_tn")),
        v(tb.get("i_r")),v(tb.get("i_s")),v(tb.get("i_t")),
        v(tb.get("i_n")),v(tb.get("i_tie")),
        v(tb.get("t_r")),v(tb.get("t_s")),v(tb.get("t_t")),
        v(tb.get("t_n")),v(tb.get("t_tie")),
        "OK" if tb.get("luz_piloto") else "—",
        "OK" if tb.get("chapas") else "—"
    ]
    fills2 = ["FFFFFF"]*16 + [("E2EFDA" if tb.get("luz_piloto") else "FFFFFF"),
                               ("E2EFDA" if tb.get("chapas")     else "FFFFFF")]
    for i,(val,fill) in enumerate(zip(vals,fills2)):
        val_cell(vrow.cells[i], val, fill)

    # Fila neutro-tierra + observaciones
    lbl4 = tbl.add_row()
    ntie_labels = ["R/TIE","S/TIE","T/TIE","N/TIE"]
    for i,lbl in enumerate(ntie_labels):
        hdr_cell(lbl4.cells[i], lbl, GR_HEX)
        lbl4.cells[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(0x1F,0x4E,0x79)
    obs_cell = lbl4.cells[4].merge(lbl4.cells[17])
    set_cell_bg(obs_cell, "FAFAFA")
    set_cell_borders(obs_cell)
    obs_cell.text = ""
    p_obs = obs_cell.paragraphs[0]
    if tb.get("obs"):
        r_lbl = p_obs.add_run("Obs: "); r_lbl.bold=True; r_lbl.font.size=Pt(8); r_lbl.font.name="Arial"
        r_obs = p_obs.add_run(tb["obs"]); r_obs.font.size=Pt(8); r_obs.font.name="Arial"
    else:
        r_no = p_obs.add_run("Sin observaciones"); r_no.italic=True; r_no.font.size=Pt(7); r_no.font.name="Arial"; r_no.font.color.rgb=RGBColor(0xAA,0xAA,0xAA)

    # Fila valores neutro-tierra
    vrow2 = tbl.add_row()
    ntie_vals = [v(tb.get("v_rtie")),v(tb.get("v_stie")),v(tb.get("v_ttie")),v(tb.get("v_ntie"))]
    for i,val in enumerate(ntie_vals):
        val_cell(vrow2.cells[i], val)
    empty = vrow2.cells[4].merge(vrow2.cells[17])
    set_cell_bg(empty,"FAFAFA"); set_cell_borders(empty)

    # Anchos
    ws = [1.4,1.4,1.4,1.4,1.4,1.4, 1.4,1.4,1.4,1.4,1.5, 1.1,1.1,1.1,1.1,1.1, 1.5,1.5]
    set_col_widths(tbl, ws)
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

# ── Sección gráficos ───────────────────────────────────────────────────────
def seccion_graficos(doc, graf_paths):
    add_heading(doc, "Análisis histórico de parámetros")
    add_para(doc, "Los siguientes gráficos muestran la evolución mensual de tensiones, corrientes y temperaturas.")
    titulos = {
        "tension":       "Evolución de tensiones (V) — últimos meses",
        "corriente":     "Evolución de corrientes (A) — últimos meses",
        "temperatura":   "Evolución de temperaturas (°C) — últimos meses",
        "barras_tension":"Tensión R/S por tablero — mes actual",
    }
    for key, titulo in titulos.items():
        path = graf_paths.get(key)
        if not path or not os.path.exists(path):
            continue
        add_para(doc, titulo, size=9, before=6, after=2)
        try:
            doc.add_picture(path, width=Cm(16))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        except Exception as e:
            add_para(doc, f"[Gráfico no disponible: {e}]", size=8)
    doc.add_page_break()

# ── Sección fotos ──────────────────────────────────────────────────────────
def seccion_fotos(doc, fotografias):
    doc.add_page_break()
    add_heading(doc, "Registro fotográfico")
    if not fotografias:
        add_para(doc, "No se adjuntaron fotografías en esta toma de parámetros.", size=9)
        return
    add_para(doc, f"Se adjuntan {len(fotografias)} fotografía(s) tomadas durante la revisión.", size=9, after=6)
    for i in range(0, len(fotografias), 2):
        tbl = doc.add_table(rows=1, cols=3)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        for j, col_idx in enumerate([0, 2]):
            if i+j >= len(fotografias):
                continue
            foto = fotografias[i+j]
            cell = tbl.rows[0].cells[col_idx]
            set_cell_borders(cell, "DDDDDD")
            set_cell_margins(cell, 60, 60, 60, 60)
            try:
                img_data = base64.b64decode(foto["data"])
                img_buf  = BytesIO(img_data)
                p_img = cell.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p_img.add_run()
                run.add_picture(img_buf, width=Cm(7.5))
            except Exception as e:
                cell_run(cell, f"[Foto no disponible]", size=8)
            p_tb = cell.add_paragraph(foto.get("tablero",""))
            p_tb.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_tb.runs[0].bold = True; p_tb.runs[0].font.size = Pt(8); p_tb.runs[0].font.name = "Arial"
            if foto.get("descripcion"):
                p_d = cell.add_paragraph(foto["descripcion"])
                p_d.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_d.runs[0].font.size = Pt(7); p_d.runs[0].font.name = "Arial"
            p_fh = cell.add_paragraph(foto.get("fecha_hora",""))
            p_fh.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_fh.runs[0].font.size = Pt(7); p_fh.runs[0].font.name = "Arial"
            p_fh.runs[0].font.color.rgb = RGBColor(0x99,0x99,0x99)
        # Separador
        sep = tbl.rows[0].cells[1]
        set_no_borders(sep)
        tbl.rows[0].cells[0].width = Cm(7.5)
        tbl.rows[0].cells[1].width = Cm(0.5)
        tbl.rows[0].cells[2].width = Cm(7.5)
        doc.add_paragraph().paragraph_format.space_after = Pt(6)

# ── Documento principal ────────────────────────────────────────────────────
def build_doc(datos):
    doc  = Document()
    sect = doc.sections[0]
    sect.page_width   = Cm(21)
    sect.page_height  = Cm(29.7)
    sect.left_margin  = sect.right_margin  = Cm(1.5)
    sect.top_margin   = sect.bottom_margin = Cm(2)

    tbs = datos.get("tableros", [])

    # ── PORTADA ────────────────────────────────────────────────────────────
    p_logo = doc.add_paragraph()
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_logo.paragraph_format.space_before = Pt(40)
    if os.path.exists(LOGO_PATH):
        run_logo = p_logo.add_run()
        run_logo.add_picture(LOGO_PATH, width=Cm(9))

    p_tit = doc.add_paragraph()
    p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tit.paragraph_format.space_before = Pt(16)
    p_tit.paragraph_format.space_after  = Pt(8)
    rt = p_tit.add_run("INFORME DE MANTENCIÓN ELÉCTRICA")
    rt.bold = True; rt.font.size = Pt(20); rt.font.name = "Arial"; rt.font.color.rgb = AZ

    p_ed = doc.add_paragraph()
    p_ed.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ed.paragraph_format.space_after = Pt(24)
    re = p_ed.add_run(datos.get("edificio","").upper())
    re.bold = True; re.font.size = Pt(14); re.font.name = "Arial"; re.font.color.rgb = AZ2

    tbl_cov = doc.add_table(rows=5, cols=2)
    tbl_cov.alignment = WD_TABLE_ALIGNMENT.CENTER
    rows_data = [
        ("Período",           f"{datos.get('mes','')} {datos.get('anio','')}"),
        ("Técnico",           datos.get("tecnico","")),
        ("Jefe de Mantención",datos.get("jefe_mantencion","—")),
        ("Fecha",             datos.get("fecha","")),
        ("Tableros revisados",str(len(tbs))),
    ]
    for i,(lbl,val) in enumerate(rows_data):
        hdr_cell(tbl_cov.rows[i].cells[0], lbl, AZ_HEX)
        val_cell(tbl_cov.rows[i].cells[1], val, "EBF3FB", bold=(i==4), size=9, align="left")
    set_col_widths(tbl_cov, [4, 8.5])
    doc.add_page_break()

    # ── RESUMEN DE ALERTAS ─────────────────────────────────────────────────
    seccion_alertas(doc, tbs)
    doc.add_page_break()

    # ── INTRODUCCIÓN ───────────────────────────────────────────────────────
    add_heading(doc, "Introducción")
    add_para(doc, f"El presente informe expone los resultados de la toma de parámetros eléctricos realizada en {datos.get('edificio','')} durante {datos.get('mes','')} de {datos.get('anio','')}.", size=9, after=6)
    add_heading(doc, "Trabajos ejecutados", level=2)
    for item in ["Limpieza periódica mensual de tableros eléctricos.",
                 "Limpieza de gabinete con solvente dieléctrico.",
                 "Inspección de componentes en tableros eléctricos.",
                 "Toma de parámetros eléctricos: voltaje, corrientes, secuencias.",
                 "Toma de temperatura de conductores y protecciones."]:
        p_b = doc.add_paragraph(style="List Bullet")
        p_b.paragraph_format.space_after = Pt(2)
        run = p_b.add_run(item)
        run.font.size = Pt(9); run.font.name = "Arial"
    doc.add_page_break()

    # ── PARÁMETROS ─────────────────────────────────────────────────────────
    add_heading(doc, "Parámetros registrados")
    for tb in tbs:
        tabla_tablero(doc, tb)
    doc.add_page_break()

    # ── GRÁFICOS ───────────────────────────────────────────────────────────
    graf_paths = datos.get("_graf_paths", {})
    seccion_graficos(doc, graf_paths)

    # ── FOTOS ──────────────────────────────────────────────────────────────
    seccion_fotos(doc, datos.get("fotografias", []))

    # ── OBSERVACIONES Y FIRMA ─────────────────────────────────────────────
    doc.add_page_break()
    add_heading(doc, "Observaciones y conclusiones")
    con_obs = [tb for tb in tbs if tb.get("obs","").strip()]
    if con_obs:
        add_para(doc, "Se registraron las siguientes observaciones:", size=9, after=4)
        for tb in con_obs:
            p_b = doc.add_paragraph(style="List Bullet")
            p_b.paragraph_format.space_after = Pt(2)
            r1 = p_b.add_run(f"Tablero {tb.get('nombre','')} (Piso {tb.get('piso','')}): ")
            r1.bold = True; r1.font.size = Pt(9); r1.font.name = "Arial"
            r2 = p_b.add_run(tb["obs"])
            r2.font.size = Pt(9); r2.font.name = "Arial"
    else:
        add_para(doc, "No se registraron observaciones relevantes. Los parámetros se encuentran dentro de los rangos normales.", size=9)

    doc.add_paragraph()
    tbl_firma = doc.add_table(rows=1, cols=2)
    tbl_firma.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(tbl_firma, [8, 8])
    for i, (titulo, nombre) in enumerate([
        ("Elaborado por:", datos.get("tecnico","___________________")),
        ("Revisado por:",  datos.get("jefe_mantencion","___________________")),
    ]):
        cell = tbl_firma.rows[0].cells[i]
        set_no_borders(cell)
        set_cell_margins(cell, 80, 80, 80, 80)
        cell.text = ""
        p1 = cell.add_paragraph()
        r = p1.add_run(titulo); r.bold=True; r.font.size=Pt(9); r.font.name="Arial"; r.font.color.rgb=AZ

        # Firma digital si existe (solo celda del técnico)
        firma_b64 = datos.get("firma_b64") if i==0 else None
        if firma_b64:
            try:
                img_data = base64.b64decode(firma_b64)
                img_buf  = BytesIO(img_data)
                p_sig = cell.add_paragraph()
                p_sig.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p_sig.runs[0].add_picture(img_buf, width=Cm(5), height=Cm(1.5)) if p_sig.runs else p_sig.add_run().add_picture(img_buf, width=Cm(5), height=Cm(1.5))
            except:
                p_ln = cell.add_paragraph("_______________________________")
                p_ln.runs[0].font.size=Pt(9); p_ln.runs[0].font.name="Arial"; p_ln.runs[0].font.color.rgb=RGBColor(0x88,0x88,0x88)
        else:
            p_ln = cell.add_paragraph("_______________________________")
            p_ln.paragraph_format.space_before = Pt(20)
            p_ln.runs[0].font.size=Pt(9); p_ln.runs[0].font.name="Arial"; p_ln.runs[0].font.color.rgb=RGBColor(0x88,0x88,0x88)

        p2 = cell.add_paragraph(nombre); p2.runs[0].bold=True; p2.runs[0].font.size=Pt(10); p2.runs[0].font.name="Arial"
        rol = "Técnico Electricista" if i==0 else "Jefe de Mantención"
        p3 = cell.add_paragraph(rol); p3.runs[0].font.size=Pt(8); p3.runs[0].font.name="Arial"; p3.runs[0].font.color.rgb=RGBColor(0x66,0x66,0x66)
        p4 = cell.add_paragraph(f"{datos.get('mes','')} {datos.get('anio','')}"); p4.runs[0].font.size=Pt(8); p4.runs[0].font.name="Arial"; p4.runs[0].font.color.rgb=RGBColor(0x66,0x66,0x66)

    return doc

# ── Main ───────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    json_path = sys.argv[1]
    out_path  = sys.argv[2]
    with open(json_path, "r", encoding="utf-8") as f:
        datos = json.load(f)
    doc = build_doc(datos)
    doc.save(out_path)
    print(f"✓ Informe guardado en: {out_path}")
