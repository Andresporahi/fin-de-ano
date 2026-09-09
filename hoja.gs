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
 * La ficha de cada hotel se guarda en la hoja Planes (solo con la clave ADMIN_CLAVE).
 * Después de pegar este archivo: Guardar → Implementar → Nueva implementación (o nueva versión).
 */

const PAX = 10;
const NOCHES = 5;
const PRESU_DEF = 18000000;
const SPREADSHEET_ID = "1hbTsqNexVURU59Mv1M85MeLXMJSkFhXcr6t4ouimy98";

function idHoja_(raw) {
  const s = String(raw || "").trim();
  const m = s.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (m) return m[1];
  return s.split("/")[0].split("?")[0].split("#")[0];
}

function libro_() {
  const id = idHoja_(SPREADSHEET_ID);
  if (id) return SpreadsheetApp.openById(id);
  const activo = SpreadsheetApp.getActiveSpreadsheet();
  if (activo) return activo;
  throw new Error("Pega el ID de la hoja en SPREADSHEET_ID (está en la URL, entre /d/ y /edit).");
}

const CATALOGO = [
  {id:"porvenir", nombre:"Hotel El Porvenir", zona:"Coveñas · Sector El Porvenir", acceso:"tierra", total:4500000, estado:"viva"},
  {id:"laguna", nombre:"Laguna Beach", zona:"Santiago de Tolú · Playa El Francés", acceso:"tierra", total:21450000, estado:"viva"},
  {id:"arena", nombre:"Hotel Arena Beach", zona:"Arroyo de Piedra · Km 22 vía al mar", acceso:"tierra", total:9562500, estado:"viva"},
  {id:"baluarte", nombre:"Baluarte Cartagena Boutique", zona:"Cartagena · Bocagrande", acceso:"tierra", total:11375000, estado:"secundaria"},
  {id:"rodadero", nombre:"Rodadero Suites", zona:"Santa Marta · El Rodadero", acceso:"tierra", total:6166667, estado:"secundaria"},
  {id:"mucura", nombre:"Múcura Club Hotel", zona:"Isla Múcura · San Bernardo", acceso:"isla", total:15610000, estado:"secundaria"},
  {id:"playita", nombre:"La Playita", zona:"Isla Fuerte · Bolívar", acceso:"isla", total:20618684, estado:"secundaria"},
  {id:"coral", nombre:"Hotel Coral de Fuego", zona:"Isla Fuerte · Playa San Diego", acceso:"isla", total:27250000, estado:"descartable"},
  {id:"mulata", nombre:"Hotel Isla Mulata", zona:"Isla Grande · Islas del Rosario", acceso:"isla", total:20381008, estado:"descartable"},
  {id:"marazao", nombre:"Marazao Beach", zona:"Ubicación por confirmar", acceso:"isla", total:21149004, estado:"descartable"},
  {id:"tintipan", nombre:"Hotel Tintipán", zona:"Isla Tintipán · San Bernardo", acceso:"isla", total:23950000, estado:"descartable"},
  {id:"river", nombre:"Hotel River City", zona:"Montería · Córdoba", acceso:"ruta", total:3250000, estado:"complemento"}
];

const TAREAS = [
  {id:"p2", prio:"alta", t:"Laguna Beach: confirmar la quinta noche (29 dic – 3 ene)", q:"Angela"},
  {id:"p3", prio:"media", t:"Laguna Beach: pedir la mezcla de suites por escrito", q:"Angela"},
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
const ADMIN_CLAVE = "FinDeAno2026";

function slug_(s) {
  return String(s || "").toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "")
    .substring(0, 40) || "votante";
}

function votoReal_(v) {
  if (!v) return false;
  return Object.keys(v.estrellas || {}).some(function (k) {
    const n = v.estrellas[k];
    return typeof n === "number" && n > 0;
  });
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

function claveOk_(p) {
  return String((p && p.clave) || "") === ADMIN_CLAVE;
}

function planes_() {
  const all = PropertiesService.getScriptProperties().getProperties();
  const out = {};
  Object.keys(all).forEach(function (k) {
    if (k.indexOf("plan_") !== 0) return;
    try { out[k.slice(5)] = JSON.parse(all[k]); } catch (e) {}
  });
  return out;
}

function textoCorto_(x, n) {
  return String(x == null ? "" : x).substring(0, n || 500);
}

function limpiarPlan_(data) {
  const si = Array.isArray(data.si) ? data.si.slice(0, 12).map(function (x) {
    return textoCorto_(x, 400);
  }).filter(Boolean) : [];
  const no = Array.isArray(data.no) ? data.no.slice(0, 12).map(function (x) {
    return textoCorto_(x, 400);
  }).filter(Boolean) : [];
  const extra = Array.isArray(data.extra) ? data.extra.slice(0, 16).map(function (p) {
    if (Array.isArray(p)) return [textoCorto_(p[0], 80), textoCorto_(p[1], 400)];
    return [textoCorto_(p, 80), ""];
  }).filter(function (p) { return p[0] || p[1]; }) : [];
  const total = Number(data.total);
  let cabe = null;
  if (data.cabe === true || data.cabe === 1 || data.cabe === "1") cabe = true;
  if (data.cabe === false || data.cabe === 0 || data.cabe === "0") cabe = false;
  const acceso = String(data.acceso || "");
  return {
    id: textoCorto_(data.id, 40),
    nombre: textoCorto_(data.nombre, 80),
    zona: textoCorto_(data.zona, 120),
    tipo: textoCorto_(data.tipo, 80),
    acceso: ["tierra", "isla", "ruta"].indexOf(acceso) >= 0 ? acceso : "tierra",
    horas: textoCorto_(data.horas, 80),
    total: isFinite(total) && total >= 0 ? total : null,
    cotizado: textoCorto_(data.cotizado, 500),
    detalleCosto: textoCorto_(data.detalleCosto, 500),
    comida: textoCorto_(data.comida, 160),
    estado: textoCorto_(data.estado, 40),
    cabe: cabe,
    fuente: textoCorto_(data.fuente, 120),
    mapsq: textoCorto_(data.mapsq, 160),
    wa: textoCorto_(data.wa, 200),
    si: si,
    no: no,
    extra: extra
  };
}

function escribirPlanes_(planes) {
  const ss = libro_();
  const sh = hoja_(ss, "Planes");
  const headers = ["id", "nombre", "zona", "tipo", "acceso", "horas", "total", "cotizado", "detalle_costo", "comida", "estado", "caben_10", "fuente", "maps", "whatsapp", "a_favor", "alertas", "letra_menuda"];
  const filas = [headers];
  CATALOGO.forEach(function (cat) {
    const o = planes[cat.id] || {};
    const si = Array.isArray(o.si) ? o.si.join("\n") : "";
    const no = Array.isArray(o.no) ? o.no.join("\n") : "";
    const extra = Array.isArray(o.extra) ? o.extra.map(function (p) {
      return (p[0] || "") + ": " + (p[1] || "");
    }).join("\n") : "";
    let cabe = "";
    if (o.cabe === true) cabe = 1;
    if (o.cabe === false) cabe = 0;
    filas.push([
      cat.id,
      o.nombre || cat.nombre || "",
      o.zona || cat.zona || "",
      o.tipo || "",
      o.acceso || cat.acceso || "",
      o.horas || "",
      o.total != null ? o.total : (cat.total != null ? cat.total : ""),
      o.cotizado || "",
      o.detalleCosto || "",
      o.comida || "",
      o.estado || cat.estado || "",
      cabe,
      o.fuente || "",
      o.mapsq || "",
      o.wa || "",
      si,
      no,
      extra
    ]);
  });
  volcar_(sh, filas);
}

function guardarPlan_(data) {
  const limpio = limpiarPlan_(data);
  const id = String(limpio.id || "").replace(/[^a-z0-9_-]/gi, "");
  if (!id) throw new Error("Falta el id del plan.");
  limpio.id = id;
  PropertiesService.getScriptProperties().setProperty("plan_" + id, JSON.stringify(limpio));
  escribirPlanes_(planes_());
}

function snapshot_(s) {
  return { ok: true, votos: s.votos, config: s.config, hechos: s.hechos, planes: planes_() };
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
    res[o.id] = {
      prom: prom, notas: notas.length,
      puntaje: notas.length ? (prom / 5) * 100 : 0
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
      if (!e) return;
      filas.push([t, v.nombre || x.id, o.id, o.nombre, e, 0, 0]);
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
    } else if (p.op === "admin") {
      if (!claveOk_(p)) return responder_(e, { ok: false, error: "clave", votos: s.votos, config: s.config, hechos: s.hechos, planes: planes_() });
      return responder_(e, snapshot_(s));
    } else if (p.op === "plan" && data && data.id) {
      if (!claveOk_(p)) return responder_(e, { ok: false, error: "clave", votos: s.votos, config: s.config, hechos: s.hechos, planes: planes_() });
      guardarPlan_(data);
      return responder_(e, snapshot_(s));
    }
  } catch (err) {
    return responder_(e, { ok: false, error: String(err), votos: s.votos, config: s.config, hechos: s.hechos, planes: planes_() });
  }
  return responder_(e, snapshot_(s));
}
