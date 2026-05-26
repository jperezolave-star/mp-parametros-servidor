const fs   = require("fs");
const path = require("path");
const { execSync } = require("child_process");
const LOGO_PATH = "/home/claude/logo_transparente.png";
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat, ImageRun
} = require("docx");

// ── Datos de ejemplo ────────────────────────────────────────────────────────
const DATOS_EJEMPLO = {
  edificio: "Torre B - Parque Titanium",
  mes: "Mayo", anio: "2026",
  tecnico: "Jose Perez Olave",
  jefe_mantencion: "Carlos Muñoz Vera",
  fecha: "13-05-2026",
  historial_meses: ["Ene","Feb","Mar","Abr","May"],
  historial_tension: [
    {rs:378,rt:380,st:382},{rs:379,rt:381,st:383},{rs:380,rt:382,st:384},
    {rs:381,rt:382,st:385},{rs:382,rt:385,st:384}
  ],
  historial_corriente: [
    {r:85,s:98,t:94},{r:87,s:100,t:96},{r:89,s:102,t:98},
    {r:91,s:104,t:99},{r:90,s:105,t:100}
  ],
  historial_temperatura: [
    {r:22,s:22,t:22},{r:23,s:23,t:23},{r:24,s:24,t:24},
    {r:25,s:25,t:25},{r:26,s:26,t:26}
  ],
  fotografias: [],
  tableros: [
    {n:1,nombre:"Piso 3",servicio:"EMERG",piso:3,v_rs:380,v_rt:382,v_st:385,v_rn:219,v_sn:220,v_tn:221,v_rtie:219,v_stie:220,v_ttie:221,v_ntie:0.2,i_r:0,i_s:0,i_t:0,i_n:4.4,i_tie:0.7,t_r:23,t_s:23,t_t:23,t_n:23,t_tie:23,luz_piloto:1,chapas:0,obs:"",fecha_hora:"13-05-2026 09:15"},
    {n:2,nombre:"Piso 5",servicio:"EMERG",piso:5,v_rs:380,v_rt:382,v_st:384,v_rn:220,v_sn:223,v_tn:222,v_rtie:220,v_stie:223,v_ttie:222,v_ntie:0.2,i_r:19.3,i_s:20.7,i_t:2.7,i_n:6.2,i_tie:0.4,t_r:28,t_s:28,t_t:28,t_n:28,t_tie:30,luz_piloto:0,chapas:0,obs:"Revisar fases",fecha_hora:"13-05-2026 09:31"},
    {n:3,nombre:"TGAFyCOMP. EMERGENCIA",servicio:"EMERG",piso:-2,v_rs:382,v_rt:385,v_st:384,v_rn:219,v_sn:220,v_tn:222,v_rtie:219,v_stie:220,v_ttie:222,v_ntie:0.5,i_r:90,i_s:105,i_t:100,i_n:42.3,i_tie:7.5,t_r:26,t_s:26,t_t:26,t_n:26,t_tie:24,luz_piloto:0,chapas:0,obs:"",fecha_hora:"13-05-2026 14:05"},
  ]
};

// ── Helpers ────────────────────────────────────────────────────────────────
const COLOR_AZUL   = "1F4E79";
const COLOR_AZUL2  = "2E75B6";
const COLOR_GRIS   = "F2F2F2";
const COLOR_VERDE  = "E2EFDA";
const COLOR_TITULO = "FFFFFF";

const borde  = (c="CCCCCC") => ({ style: BorderStyle.SINGLE, size:1, color:c });
const bordes = (c="CCCCCC") => ({ top:borde(c),bottom:borde(c),left:borde(c),right:borde(c) });
const sinBorde = () => ({ style:BorderStyle.NONE, size:0, color:"FFFFFF" });
const sinBordes = () => ({ top:sinBorde(),bottom:sinBorde(),left:sinBorde(),right:sinBorde() });

function txt(text, opts={}) {
  return new TextRun({ text:String(text??""), font:"Arial", size:opts.size||20,
    bold:opts.bold||false, color:opts.color||"000000", italics:opts.italics||false, ...opts });
}
function par(children, opts={}) {
  return new Paragraph({
    alignment: opts.align||AlignmentType.LEFT,
    spacing: { before:opts.before||0, after:opts.after||120 },
    children: Array.isArray(children)?children:[children],
  });
}
function logoRun(wCm, hCm) {
  const EMU=914400, CM=EMU/2.54;
  return new ImageRun({ data:fs.readFileSync(LOGO_PATH),
    transformation:{width:Math.round(wCm*CM),height:Math.round(hCm*CM)}, type:"png" });
}
function imgRun(buf, wCm, hCm) {
  const EMU=914400, CM=EMU/2.54;
  return new ImageRun({ data:buf,
    transformation:{width:Math.round(wCm*CM),height:Math.round(hCm*CM)}, type:"png" });
}
function celdaH(texto, ancho, color=COLOR_AZUL2) {
  return new TableCell({ width:{size:ancho,type:WidthType.DXA}, borders:bordes("FFFFFF"),
    shading:{fill:color,type:ShadingType.CLEAR}, verticalAlign:VerticalAlign.CENTER,
    margins:{top:60,bottom:60,left:80,right:80},
    children:[par([txt(texto,{bold:true,color:COLOR_TITULO,size:16})],{align:AlignmentType.CENTER})] });
}
function celda(texto, ancho, opts={}) {
  return new TableCell({ width:{size:ancho,type:WidthType.DXA}, borders:bordes(opts.borderColor||"CCCCCC"),
    shading:{fill:opts.fill||"FFFFFF",type:ShadingType.CLEAR}, verticalAlign:VerticalAlign.CENTER,
    margins:{top:50,bottom:50,left:80,right:80},
    children:[par([txt(texto,{size:opts.size||17,bold:opts.bold||false,color:opts.color||"000000"})],{align:opts.align||AlignmentType.CENTER})] });
}
const v = val => (val===null||val===undefined)?"—":String(val);

// ── Tabla de un tablero ────────────────────────────────────────────────────
function tablaTablero(tb) {
  const W=9360;
  const hRow = new TableRow({ children:[
    new TableCell({ columnSpan:18, width:{size:W,type:WidthType.DXA},
      borders:bordes("1F4E79"), shading:{fill:COLOR_AZUL,type:ShadingType.CLEAR},
      margins:{top:80,bottom:80,left:120,right:120},
      children:[par([
        txt(`N° ${tb.n}  —  ${tb.nombre}`,{bold:true,color:"FFFFFF",size:18}),
        txt(`   Piso ${tb.piso}   |   ${tb.servicio}`,{color:"BDD7EE",size:16}),
        tb.fecha_hora?txt(`   ·   ${tb.fecha_hora}`,{color:"9DC3E6",size:15,italics:true}):txt(""),
      ])]})
  ]});
  const subH = new TableRow({ children:[
    celdaH("TENSIONES (V)",3600,COLOR_AZUL2), celdaH("",100,COLOR_AZUL2),
    celdaH("CORRIENTES (A)",2200,"1F6B3E"), celdaH("",100,"1F6B3E"),
    celdaH("TEMPERATURAS (°C)",2200,"7F6000"), celdaH("ESTADO",1160,"843C0C"),
  ]});
  const lbl1 = new TableRow({ children:[
    celda("R/S",540,{fill:COLOR_GRIS,bold:true}),celda("R/T",540,{fill:COLOR_GRIS,bold:true}),
    celda("S/T",540,{fill:COLOR_GRIS,bold:true}),celda("R/N",540,{fill:COLOR_GRIS,bold:true}),
    celda("S/N",540,{fill:COLOR_GRIS,bold:true}),celda("T/N",540,{fill:COLOR_GRIS,bold:true}),
    celda("CTE R",540,{fill:COLOR_VERDE,bold:true}),celda("CTE S",540,{fill:COLOR_VERDE,bold:true}),
    celda("CTE T",540,{fill:COLOR_VERDE,bold:true}),celda("CTE N",540,{fill:COLOR_VERDE,bold:true}),
    celda("CTE TIE",580,{fill:COLOR_VERDE,bold:true}),
    celda("Tº R",440,{fill:"FFF2CC",bold:true}),celda("Tº S",440,{fill:"FFF2CC",bold:true}),
    celda("Tº T",440,{fill:"FFF2CC",bold:true}),celda("Tº N",440,{fill:"FFF2CC",bold:true}),
    celda("Tº TIE",440,{fill:"FFF2CC",bold:true}),
    celda("Piloto",580,{fill:"FCE4D6",bold:true}),celda("Chapas",580,{fill:"FCE4D6",bold:true}),
  ]});
  const val1 = new TableRow({ children:[
    celda(v(tb.v_rs),540),celda(v(tb.v_rt),540),celda(v(tb.v_st),540),
    celda(v(tb.v_rn),540),celda(v(tb.v_sn),540),celda(v(tb.v_tn),540),
    celda(v(tb.i_r),540),celda(v(tb.i_s),540),celda(v(tb.i_t),540),
    celda(v(tb.i_n),540),celda(v(tb.i_tie),580),
    celda(v(tb.t_r),440),celda(v(tb.t_s),440),celda(v(tb.t_t),440),
    celda(v(tb.t_n),440),celda(v(tb.t_tie),440),
    celda(tb.luz_piloto?"OK":"—",580,{fill:tb.luz_piloto?"E2EFDA":"FFFFFF"}),
    celda(tb.chapas?"OK":"—",580,{fill:tb.chapas?"E2EFDA":"FFFFFF"}),
  ]});
  const lbl2 = new TableRow({ children:[
    celda("R/TIE",540,{fill:COLOR_GRIS,bold:true}),celda("S/TIE",540,{fill:COLOR_GRIS,bold:true}),
    celda("T/TIE",540,{fill:COLOR_GRIS,bold:true}),celda("NEU/TIE",540,{fill:COLOR_GRIS,bold:true}),
    new TableCell({ columnSpan:14, width:{size:5660,type:WidthType.DXA},
      borders:bordes("CCCCCC"), shading:{fill:"FAFAFA",type:ShadingType.CLEAR},
      margins:{top:50,bottom:50,left:80,right:80},
      children:[par(tb.obs?[txt("Obs: ",{bold:true,size:16}),txt(tb.obs,{size:16,color:"4A4A4A"})]:[txt("Sin observaciones",{size:15,color:"AAAAAA",italics:true})])]})
  ]});
  const val2 = new TableRow({ children:[
    celda(v(tb.v_rtie),540),celda(v(tb.v_stie),540),
    celda(v(tb.v_ttie),540),celda(v(tb.v_ntie),540),
    new TableCell({ columnSpan:14, width:{size:5660,type:WidthType.DXA},
      borders:bordes("CCCCCC"), shading:{fill:"FAFAFA",type:ShadingType.CLEAR},
      margins:{top:50,bottom:50,left:80,right:80},
      children:[par([txt("")])]})
  ]});
  return new Table({ width:{size:W,type:WidthType.DXA},
    columnWidths:[540,540,540,540,540,540,540,540,540,540,580,440,440,440,440,440,580,580],
    rows:[hRow,subH,lbl1,val1,lbl2,val2] });
}

// ── Sección de gráficos ────────────────────────────────────────────────────
function seccionGraficos(grafPaths) {
  const W = 9360, EM = 914400, CM = EM/2.54;
  const items = [];

  items.push(new Paragraph({
    spacing:{before:240,after:160},
    children:[txt("Análisis histórico de parámetros",{bold:true,size:28,color:COLOR_AZUL})]
  }));
  items.push(par([txt("Los siguientes gráficos muestran la evolución mensual de tensiones, corrientes y temperaturas registradas en el edificio, permitiendo identificar tendencias y variaciones en el tiempo.",{size:20})],{after:200}));

  const titulos = {
    tension:      "Evolución de tensiones (V) — últimos meses",
    corriente:    "Evolución de corrientes (A) — últimos meses",
    temperatura:  "Evolución de temperaturas (°C) — últimos meses",
    barras_tension: "Tensión R/S por tablero — mes actual",
  };

  for (const [key, titulo] of Object.entries(titulos)) {
    if (!grafPaths[key] || !fs.existsSync(grafPaths[key])) continue;
    const buf = fs.readFileSync(grafPaths[key]);
    // calc height: gráficos tienen proporción ~7.2:3 (width:height)
    const wCm = 16.5, hCm = +(wCm * 3/7.2).toFixed(2);
    items.push(par([txt(titulo,{bold:true,size:20,color:COLOR_AZUL2})],{before:160,after:80}));
    items.push(new Paragraph({
      spacing:{before:0,after:200},
      alignment:AlignmentType.CENTER,
      children:[imgRun(buf,wCm,hCm)]
    }));
  }
  return items;
}

// ── Sección de fotografías ─────────────────────────────────────────────────
function seccionFotos(fotografias) {
  const items = [];
  items.push(new Paragraph({ children:[new PageBreak()], spacing:{before:0,after:0} }));
  items.push(new Paragraph({
    spacing:{before:240,after:160},
    children:[txt("Registro fotográfico",{bold:true,size:28,color:COLOR_AZUL})]
  }));

  if (!fotografias || !fotografias.length) {
    items.push(par([txt("No se adjuntaron fotografías en esta toma de parámetros.",{size:20,color:"888888",italics:true})],{after:200}));
    return items;
  }

  items.push(par([txt(`Se adjuntan ${fotografias.length} fotografía(s) tomadas durante la revisión mensual.`,{size:20})],{after:200}));

  // Agrupar en pares para tabla 2 columnas
  for (let i=0; i<fotografias.length; i+=2) {
    const fotoA = fotografias[i];
    const fotoB = fotografias[i+1] || null;

    const makeCell = (foto) => {
      if (!foto) return new TableCell({
        width:{size:4560,type:WidthType.DXA},
        borders:sinBordes(),
        children:[par([txt("")])]
      });

      const buf = Buffer.isBuffer(foto.data) ? foto.data : Buffer.from(foto.data,'base64');
      const EM=914400, CM=EM/2.54, wCm=7.5, hCm=5.6;
      return new TableCell({
        width:{size:4560,type:WidthType.DXA},
        borders:bordes("DDDDDD"),
        shading:{fill:"FAFAFA",type:ShadingType.CLEAR},
        margins:{top:120,bottom:120,left:120,right:120},
        children:[
          new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:80},
            children:[imgRun(buf,wCm,hCm)]}),
          par([txt(foto.tablero||"",{bold:true,size:17})],{align:AlignmentType.CENTER,after:20}),
          par([txt(foto.descripcion||"",{size:16,color:"666666",italics:true})],{align:AlignmentType.CENTER,after:40}),
          par([txt(foto.fecha_hora||"",{size:14,color:"999999"})],{align:AlignmentType.CENTER,after:0}),
        ]
      });
    };

    items.push(new Table({
      width:{size:9360,type:WidthType.DXA},
      columnWidths:[4560,240,4560],
      rows:[new TableRow({children:[makeCell(fotoA), new TableCell({width:{size:240,type:WidthType.DXA},borders:sinBordes(),children:[par([txt("")])]}), makeCell(fotoB)]})]
    }));
    items.push(par([txt("")],{after:160}));
  }
  return items;
}

// ── Documento ──────────────────────────────────────────────────────────────
async function buildDoc(datos, grafPaths) {
  const tbs = datos.tableros;
  const conObs = tbs.filter(t=>t.obs&&t.obs.trim());
  const totalT = tbs.length;

  // Portada
  const portada = [
    new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:800,after:400},children:[logoRun(10,2.27)]}),
    new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:200,after:200},
      border:{bottom:{style:BorderStyle.SINGLE,size:8,color:COLOR_AZUL2,space:1}},
      children:[txt("INFORME DE MANTENCIÓN ELÉCTRICA",{bold:true,size:40,color:COLOR_AZUL})]}),
    par([txt("")],{after:200}),
    new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:600},
      children:[txt(datos.edificio.toUpperCase(),{bold:true,size:32,color:COLOR_AZUL2})]}),
    new Table({width:{size:6000,type:WidthType.DXA},columnWidths:[2400,3600],rows:[
      new TableRow({children:[celda("Período",2400,{fill:COLOR_AZUL,bold:true,color:"FFFFFF",borderColor:COLOR_AZUL2}),celda(`${datos.mes} ${datos.anio}`,3600,{fill:"EBF3FB",bold:true,size:20,borderColor:COLOR_AZUL2})]}),
      new TableRow({children:[celda("Técnico",2400,{fill:COLOR_AZUL,bold:true,color:"FFFFFF",borderColor:COLOR_AZUL2}),celda(datos.tecnico,3600,{fill:"EBF3FB",borderColor:COLOR_AZUL2})]}),
      new TableRow({children:[celda("Jefe de Mantención",2400,{fill:COLOR_AZUL,bold:true,color:"FFFFFF",borderColor:COLOR_AZUL2}),celda(datos.jefe_mantencion||"—",3600,{fill:"EBF3FB",borderColor:COLOR_AZUL2})]}),
      new TableRow({children:[celda("Fecha",2400,{fill:COLOR_AZUL,bold:true,color:"FFFFFF",borderColor:COLOR_AZUL2}),celda(datos.fecha,3600,{fill:"EBF3FB",borderColor:COLOR_AZUL2})]}),
      new TableRow({children:[celda("Tableros revisados",2400,{fill:COLOR_AZUL,bold:true,color:"FFFFFF",borderColor:COLOR_AZUL2}),celda(String(totalT),3600,{fill:"EBF3FB",bold:true,size:22,borderColor:COLOR_AZUL2})]}),
    ]}),
    par([txt("")],{after:1000}),
    new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:0},
      border:{top:{style:BorderStyle.SINGLE,size:4,color:COLOR_AZUL2,space:1}},
      children:[txt("Toma de Parámetros Eléctricos — Documento de uso interno",{size:16,color:"888888",italics:true})]}),
    new Paragraph({children:[new PageBreak()],spacing:{before:0,after:0}}),
  ];

  // Introducción
  const intro = [
    new Paragraph({spacing:{before:240,after:160},children:[txt("Introducción",{bold:true,size:28,color:COLOR_AZUL})]}),
    par([txt(`El presente informe tiene como finalidad exponer los resultados de la toma de parámetros eléctricos realizada en ${datos.edificio} durante el mes de ${datos.mes} de ${datos.anio}.`,{size:20})],{after:200}),
    new Paragraph({spacing:{before:200,after:120},children:[txt("Trabajos ejecutados",{bold:true,size:24,color:COLOR_AZUL2})]}),
  ];
  ["Limpieza periódica mensual de tableros eléctricos.","Limpieza periódica mensual de gabinete eléctrico con solvente dieléctrico.","Inspección a componentes en tableros eléctricos.","Toma de parámetros eléctricos: voltaje, corrientes, secuencias.","Toma de temperatura de conductores y protecciones."]
    .forEach(t=>intro.push(new Paragraph({numbering:{reference:"bullets",level:0},spacing:{before:40,after:40},children:[txt(t,{size:20})]})));
  intro.push(new Paragraph({children:[new PageBreak()],spacing:{before:0,after:0}}));

  // Parámetros
  const parametros = [
    new Paragraph({spacing:{before:240,after:160},children:[txt("Parámetros registrados",{bold:true,size:28,color:COLOR_AZUL})]}),
  ];
  tbs.forEach((tb,i)=>{
    parametros.push(tablaTablero(tb));
    parametros.push(par([txt("")],{after: i<tbs.length-1?200:0}));
  });
  parametros.push(new Paragraph({children:[new PageBreak()],spacing:{before:0,after:0}}));

  // Gráficos históricos
  const grafSection = seccionGraficos(grafPaths);
  grafSection.push(new Paragraph({children:[new PageBreak()],spacing:{before:0,after:0}}));

  // Fotografías
  const fotosSection = seccionFotos(datos.fotografias||[]);

  // Observaciones y firma
  const resumen = [
    new Paragraph({children:[new PageBreak()],spacing:{before:0,after:0}}),
    new Paragraph({spacing:{before:240,after:160},children:[txt("Observaciones y conclusiones",{bold:true,size:28,color:COLOR_AZUL})]}),
  ];
  if (conObs.length) {
    resumen.push(par([txt("Se registraron las siguientes observaciones:",{size:20})],{after:160}));
    conObs.forEach(tb=>resumen.push(new Paragraph({numbering:{reference:"bullets",level:0},spacing:{before:60,after:60},
      children:[txt(`Tablero ${tb.nombre} (Piso ${tb.piso}): `,{bold:true,size:19}),txt(tb.obs,{size:19})]})));
  } else {
    resumen.push(par([txt("No se registraron observaciones relevantes. Los parámetros se encuentran dentro de los rangos normales.",{size:20})],{after:160}));
  }
  resumen.push(par([txt("")],{after:300}));
  resumen.push(new Paragraph({spacing:{before:0,after:120},border:{top:{style:BorderStyle.SINGLE,size:4,color:COLOR_AZUL2,space:4}},children:[txt("")]}));
  resumen.push(new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[4680,4680],rows:[
    new TableRow({children:[
      new TableCell({width:{size:4680,type:WidthType.DXA},borders:sinBordes(),margins:{top:100,bottom:100,left:120,right:120},children:[
        par([txt("Elaborado por:",{bold:true,size:18,color:COLOR_AZUL})],{after:400}),
        par([txt("_______________________________",{size:18,color:"888888"})],{after:60}),
        par([txt(datos.tecnico,{bold:true,size:20})],{after:40}),
        par([txt("Técnico Electricista",{size:18,color:"666666"})],{after:40}),
        par([txt(`${datos.mes} ${datos.anio}`,{size:18,color:"666666"})],{after:0}),
      ]}),
      new TableCell({width:{size:4680,type:WidthType.DXA},borders:sinBordes(),margins:{top:100,bottom:100,left:120,right:120},children:[
        par([txt("Revisado por:",{bold:true,size:18,color:COLOR_AZUL})],{after:400}),
        par([txt("_______________________________",{size:18,color:"888888"})],{after:60}),
        par([txt(datos.jefe_mantencion||"________________________",{bold:true,size:20})],{after:40}),
        par([txt("Jefe de Mantención",{size:18,color:"666666"})],{after:40}),
        par([txt(`${datos.mes} ${datos.anio}`,{size:18,color:"666666"})],{after:0}),
      ]}),
    ]})
  ]}));

  // Header / Footer
  const header = new Header({children:[
    new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[2200,4300,2860],rows:[
      new TableRow({children:[
        new TableCell({width:{size:2200,type:WidthType.DXA},borders:{top:{style:BorderStyle.NONE},bottom:borde(COLOR_AZUL2),left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},margins:{top:40,bottom:40,left:0,right:120},verticalAlign:VerticalAlign.CENTER,
          children:[new Paragraph({alignment:AlignmentType.LEFT,children:[logoRun(3.8,0.86)]})]}),
        new TableCell({width:{size:4300,type:WidthType.DXA},borders:{top:{style:BorderStyle.NONE},bottom:borde(COLOR_AZUL2),left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},margins:{top:60,bottom:60,left:120,right:120},
          children:[new Paragraph({children:[txt(datos.edificio,{bold:true,size:18,color:COLOR_AZUL})]}),new Paragraph({children:[txt(`Informe Toma de Parámetros — ${datos.mes} ${datos.anio}`,{size:15,color:"888888"})]}),]}),
        new TableCell({width:{size:2860,type:WidthType.DXA},borders:{top:{style:BorderStyle.NONE},bottom:borde(COLOR_AZUL2),left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},margins:{top:60,bottom:60,left:120,right:0},verticalAlign:VerticalAlign.CENTER,
          children:[new Paragraph({alignment:AlignmentType.RIGHT,children:[txt("Mantención Eléctrica",{size:15,color:"888888",italics:true})]})]}),
      ]})
    ]})]});

  const footer = new Footer({children:[
    new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[6000,3360],rows:[
      new TableRow({children:[
        new TableCell({width:{size:6000,type:WidthType.DXA},borders:{top:borde(COLOR_AZUL2),bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},margins:{top:60,bottom:0,left:0,right:0},
          children:[new Paragraph({children:[txt(`Técnico: ${datos.tecnico}   |   ${datos.mes} ${datos.anio}`,{size:15,color:"888888"})]}),]}),
        new TableCell({width:{size:3360,type:WidthType.DXA},borders:{top:borde(COLOR_AZUL2),bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},margins:{top:60,bottom:0,left:0,right:0},
          children:[new Paragraph({alignment:AlignmentType.RIGHT,children:[txt("Página ",{size:15,color:"888888"}),new TextRun({children:[PageNumber.CURRENT],font:"Arial",size:15,color:"888888"}),txt(" de ",{size:15,color:"888888"}),new TextRun({children:[PageNumber.TOTAL_PAGES],font:"Arial",size:15,color:"888888"})]})]}),
      ]})
    ]})]});

  const doc = new Document({
    numbering:{config:[{reference:"bullets",levels:[{level:0,format:LevelFormat.BULLET,text:"•",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}]}]},
    styles:{default:{document:{run:{font:"Arial",size:20}}}},
    sections:[{
      properties:{page:{size:{width:11906,height:16838},margin:{top:1000,right:1000,bottom:1000,left:1000}}},
      headers:{default:header},
      footers:{default:footer},
      children:[...portada,...intro,...parametros,...grafSection,...fotosSection,...resumen]
    }]
  });
  return doc;
}

// ── Main ──────────────────────────────────────────────────────────────────
async function main() {
  const datos = process.argv[2] ? JSON.parse(fs.readFileSync(process.argv[2],"utf8")) : DATOS_EJEMPLO;

  // 1. Generar gráficos con Python
  console.log("Generando gráficos...");
  const grafDir = "/tmp/graficos_informe";
  const datosFile = "/tmp/datos_informe.json";
  fs.writeFileSync(datosFile, JSON.stringify(datos));
  const result = execSync(`python3 /home/claude/generar_graficos.py ${grafDir} < ${datosFile}`, {maxBuffer:10*1024*1024});
  const grafPaths = JSON.parse(result.toString().trim());
  console.log("Gráficos:", Object.keys(grafPaths));

  // 2. Generar Word
  console.log("Generando informe Word...");
  const doc = await buildDoc(datos, grafPaths);
  const mes  = datos.mes.toUpperCase();
  const anio = datos.anio;
  const slug = datos.edificio.replace(/\s+/g,'_').replace(/[^A-Za-z0-9_]/g,'').slice(0,20);
  const out  = `/mnt/user-data/outputs/INFORME_${slug}_${mes}_${anio}.docx`;
  fs.writeFileSync(out, await Packer.toBuffer(doc));
  console.log("Informe guardado en:", out);
}

main().catch(e=>{console.error(e);process.exit(1);});
