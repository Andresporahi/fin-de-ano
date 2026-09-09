/**
 * Familia Fin de Año — Excel compartido
 *
 * NO uses Extensiones dentro de la hoja (en tu cuenta eso abre un error de Drive).
 *
 * 1. Cierra Chrome, ábrelo en incógnito, entra con UNA sola cuenta (la de la carpeta).
 * 2. Abre la hoja (icono verde, sin .xlsx) y copia la URL.
 *    Es algo como: https://docs.google.com/spreadsheets/d/PEGAR_ID_AQUI/edit
 * 3. Ve a https://script.google.com  (no pases por Extensiones).
 * 4. Nuevo proyecto → pega este archivo → pon el ID en SPREADSHEET_ID abajo.
 * 5. Guardar → Implementar → Aplicación web → Yo / Cualquiera → copia la URL /exec.
 *
 * Cada voto reescribe desde la fila 5 de la hoja Votos (mismas columnas del libro).
 * No toca Resultados ni Resumen decisión: esas hojas siguen calculando solas.
 */

const PAX = 10;
const NOCHES = 5;
const PRESU_DEF = 18000000;
const SPREADSHEET_ID = "";

function libro_() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  const activo = SpreadsheetApp.getActiveSpreadsheet();
  if (activo) return activo;
  throw new Error("Pega el ID de la hoja en SPREADSHEET_ID (está en la URL, entre /d/ y /edit).");
}

const CATALOGO = [
  {id:"porvenir", nombre:"Hotel El Porvenir", zona:"Coveñas · Sector El Porvenir", acceso:"tierra", total:4500000, estado:"viva"},
  {id:"laguna", nombre:"Laguna Beach", zona:"Santiago de Tolú · Playa El Francés", acceso:"tierra", total:null, estado:"viva"},
  {id:"arena", nombre:"Hotel Arena Beach", zona:"Arroyo de Piedra · Km 22 vía al mar", acceso:"tierra", total:9842500, estado:"viva"},
  {id:"baluarte", nombre:"Baluarte Cartagena Boutique", zona:"Cartagena · Bocagrande", acceso:"tierra", total:11375000, estado:"secundaria"},
  {id:"rodadero", nombre:"Rodadero Suites", zona:"Santa Marta · El Rodadero", acceso:"tierra", total:6166667, estado:"secundaria"},
  {id:"mucura", nombre:"Múcura Club Hotel", zona:"Isla Múcura · San Bernardo", acceso:"isla", total:15610000, estado:"secundaria"},
  {id:"playita", nombre:"La Playita", zona:"Isla Fuerte · Bolívar", acceso:"isla", total:20618684, estado:"secundaria"},
  {id:"coral", nombre:"Hotel Coral de Fuego", zona:"Isla Fuerte · Playa San Diego", acceso:"isla", total:27400000, estado:"descartable"},
  {id:"mulata", nombre:"Hotel Isla Mulata", zona:"Isla Grande · Islas del Rosario", acceso:"isla", total:20381008, estado:"descartable"},
  {id:"marazao", nombre:"Marazao Beach", zona:"Ubicación por confirmar", acceso:"isla", total:21239004, estado:"descartable"},
  {id:"tintipan", nombre:"Hotel Tintipán", zona:"Isla Tintipán · San Bernardo", acceso:"isla", total:23950000, estado:"descartada"},
  {id:"river", nombre:"Hotel River City", zona:"Montería · Córdoba", acceso:"ruta", total:3250000, estado:"complemento"}
];

const TAREAS = [
  {id:"p1", prio:"alta", t:"Confirmar por escrito que aceptan a Alma (14 kg)", q:"Angela"},
  {id:"p2", prio:"alta", t:"Laguna Beach: negociar el kilo de más", q:"Jorge"},
  {id:"p3", prio:"alta", t:"Pedir tarifa de Laguna Beach para 5 suites, 29 dic – 3 ene", q:"Angela"},
  {id:"p4", prio:"alta", t:"Averiguar el nombre real del hotel de El Porvenir", q:"Jorge"},
  {id:"p5", prio:"alta", t:"Recotizar Arena Beach por 5 noches desde el 29", q:"Angela"},
  {id:"p6", prio:"alta", t:"Confirmar si la lancha es realmente excluyente", q:"Familia"},
  {id:"p7", prio:"media", t:"Calcular la comida en Arena Beach", q:"Familia"},
  {id:"p8", prio:"media", t:"Múcura: pedir la cotización con impuestos y con la niña de 4 años", q:"Angela"},
  {id:"p9", prio:"media", t:"Definir distribución de camas y privacidad", q:"Familia"},
  {id:"p10", prio:"media", t:"Confirmar parqueadero y seguridad del vehículo", q:"Jorge"},
  {id:"p11", prio:"media", t:"Pedir política de cancelación por escrito en cada finalista", q:"Angela"},
  {id:"p12", prio:"media", t:"Reservar la parada de Montería para ida y regreso", q:"Angela"}
];

const NOMBRES_DEF = ["Jorge","Angela","Willy","Sol","Daniel","Leidy","Nelson","Alba","Gustavo","Susana"];

function slug_(s) {
  return String(s || "").toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "")
    .substring(0, 40) || "votante";
}

function votoReal_(v) {
  if (!v) return false;
  const estrellas = Object.keys(v.estrellas || {}).some(function (k) {
    const n = v.estrellas[k];
    return typeof n === "number" && n > 0;
  });
  return !!(estrellas || v.joya || v.veto);
}

function estado_() {
  const raw = PropertiesService.getScriptProperties().getProperty("estado");
  if (raw) {
    try { return JSON.parse(raw); } catch (e) {}
  }
  return {
    votos: {},
    config: { revelar: true, presupuesto: PRESU_DEF, nombres: NOMBRES_DEF.slice() },
    hechos: {}
  };
}

function guardar_(s) {
  PropertiesService.getScriptProperties().setProperty("estado", JSON.stringify(s));
  escribirHojas_(s);
}

function ts_(ms) {
  const d = new Date(ms || Date.now());
  const p = function (n) { return ("0" + n).slice(-2); };
  return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate()) + " " + p(d.getHours()) + ":" + p(d.getMinutes());
}

function hoja_(ss, nombre) {
  let sh = ss.getSheetByName(nombre);
  if (!sh) sh = ss.insertSheet(nombre);
  return sh;
}

function volcar_(sh, filas) {
  sh.clearContents();
  if (!filas.length) return;
  const cols = filas.reduce(function (m, f) { return Math.max(m, f.length); }, 0);
  const rect = filas.map(function (f) {
    const r = f.slice();
    while (r.length < cols) r.push("");
    return r.map(function (v) { return v === null || v === undefined ? "" : v; });
  });
  sh.getRange(1, 1, rect.length, cols).setValues(rect);
  sh.setFrozenRows(1);
}

function puntajes_(s) {
  const lista = Object.keys(s.votos).map(function (k) { return s.votos[k]; }).filter(votoReal_);
  const nv = lista.length;
  const res = {};
  CATALOGO.forEach(function (o) {
    const notas = lista.map(function (v) { return v.estrellas && v.estrellas[o.id]; }).filter(function (n) {
      return typeof n === "number" && n > 0;
    });
    const prom = notas.length ? notas.reduce(function (a, b) { return a + b; }, 0) / notas.length : 0;
    const joyas = lista.filter(function (v) { return v.joya === o.id; }).length;
    const vetos = lista.filter(function (v) { return v.veto === o.id; }).length;
    const base = notas.length ? (prom / 5) * 70 : 0;
    const bonus = nv ? (joyas / nv) * 30 : 0;
    res[o.id] = {
      prom: prom, notas: notas.length, joyas: joyas, vetos: vetos,
      puntaje: Math.max(0, base + bonus - vetos * 15)
    };
  });
  return { res: res, nv: nv };
}

function escribirHojas_(s) {
  const ss = libro_();
  const sh = hoja_(ss, "Votos");
  const reales = Object.keys(s.votos).map(function (id) {
    return { id: id, v: s.votos[id] };
  }).filter(function (x) { return votoReal_(x.v); });

  if (sh.getLastRow() < 4) {
    sh.getRange(1, 1).setValue("VOTOS — base de datos de la votación familiar");
    sh.getRange(2, 1).setValue("Una fila por votante y por opción. El tablero escribe desde la fila 5. No borre las filas 1 a 4.");
    sh.getRange(4, 1, 1, 7).setValues([["Fecha y hora", "Votante", "ID opción", "Opción", "Estrellas (1-5)", "Joya (1/0)", "Veto (1/0)"]]);
  }

  if (sh.getLastRow() >= 5) {
    sh.getRange(5, 1, sh.getLastRow() - 4, 7).clearContent();
  }

  const filas = [];
  reales.forEach(function (x) {
    const v = x.v;
    const t = ts_(v.ts);
    CATALOGO.forEach(function (o) {
      const e = (v.estrellas && v.estrellas[o.id]) || 0;
      const j = v.joya === o.id ? 1 : 0;
      const k = v.veto === o.id ? 1 : 0;
      if (!e && !j && !k) return;
      filas.push([t, v.nombre || x.id, o.id, o.nombre, e, j, k]);
    });
  });
  if (filas.length) sh.getRange(5, 1, filas.length, 7).setValues(filas);
}

function responder_(e, payload) {
  const body = JSON.stringify(payload);
  const cb = e && e.parameter && e.parameter.callback;
  if (cb && /^[A-Za-z_][A-Za-z0-9_]*$/.test(cb)) {
    return ContentService.createTextOutput(cb + "(" + body + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const s = estado_();
  let data = p.data;
  if (typeof data === "string" && data) {
    try { data = JSON.parse(data); } catch (err) { data = null; }
  }
  try {
    if (p.op === "voto" && data && data.nombre) {
      const id = slug_(data.nombre);
      data.ts = Date.now();
      s.votos[id] = data;
      guardar_(s);
    } else if (p.op === "borrar" && p.id) {
      delete s.votos[p.id];
      guardar_(s);
    } else if (p.op === "config" && data && typeof data === "object") {
      s.config = {
        revelar: data.revelar !== false,
        presupuesto: typeof data.presupuesto === "number" ? data.presupuesto : s.config.presupuesto,
        nombres: Array.isArray(data.nombres) && data.nombres.length ? data.nombres : s.config.nombres
      };
      guardar_(s);
    } else if (p.op === "hechos" && data && typeof data === "object") {
      s.hechos = data;
      guardar_(s);
    }
  } catch (err) {
    return responder_(e, { ok: false, error: String(err), votos: s.votos, config: s.config, hechos: s.hechos });
  }
  return responder_(e, { ok: true, votos: s.votos, config: s.config, hechos: s.hechos });
}
