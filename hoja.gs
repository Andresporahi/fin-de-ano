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
 * Cada voto reescribe la hoja Votos: detalle por persona/plan y sumatoria de estrellas.
 * Sin columnas de joya ni veto. No toca Resultados ni Resumen decisión.
 * La ficha de cada hotel se guarda en la hoja Planes (solo con la clave ADMIN_CLAVE).
 * Después de pegar este archivo: Guardar → Implementar → Nueva implementación (hace falta para guardar fichas largas).
 */

const PAX = 10;
const NOCHES = 4;
const PRESU_DEF = 18000000;
const SPREADSHEET_ID = "1SMBdOflN67BmDHENLx1Jb82sPYEqC_Rz9IUvFXGoUMg";

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
  {id:"porvenir", nombre:"Hotel El Porvenir", zona:"Coveñas · Sector El Porvenir", acceso:"tierra", total:3600000, estado:"viva", web:"https://www.booking.com/searchresults.es.html?ss=El+Porvenir+Cove%C3%B1as"},
  {id:"laguna", nombre:"Laguna Beach", zona:"Santiago de Tolú · Playa El Francés", acceso:"tierra", total:21450000, estado:"viva", web:"https://www.booking.com/searchresults.es.html?ss=Laguna+Beach+Ecohotel+Playa+El+Franc%C3%A9s+Tol%C3%BA"},
  {id:"arena", nombre:"Hotel Arena Beach", zona:"Arroyo de Piedra · Km 22 vía al mar", acceso:"tierra", total:7650000, estado:"viva", web:"https://arenabeach.co/"},
  {id:"baluarte", nombre:"Baluarte Cartagena Boutique", zona:"Cartagena · Bocagrande", acceso:"tierra", total:9100000, estado:"secundaria", web:"https://www.booking.com/hotel/co/san-pietro.es.html"},
  {id:"rodadero", nombre:"Rodadero Suites", zona:"Santa Marta · El Rodadero", acceso:"tierra", total:4933333, estado:"secundaria", web:"https://www.google.com/maps/search/?api=1&query=SGH+Rodadero+Suites+Calle+7+El+Rodadero+Santa+Marta"},
  {id:"mucura", nombre:"Múcura Club Hotel", zona:"Isla Múcura · San Bernardo", acceso:"isla", total:12488000, estado:"secundaria", web:"https://www.mucuraclubhotel.com/"},
  {id:"playita", nombre:"La Playita", zona:"Isla Fuerte · Bolívar", acceso:"isla", total:16494947, estado:"secundaria", web:"https://www.booking.com/hotel/co/la-playita-isla-fuerte.es.html"},
  {id:"coral", nombre:"Hotel Coral de Fuego", zona:"Isla Fuerte · Playa San Diego", acceso:"isla", total:21800000, estado:"descartable", web:"https://www.booking.com/hotel/co/coral-de-fuego.es.html"},
  {id:"mulata", nombre:"Hotel Isla Mulata", zona:"Isla Grande · Islas del Rosario", acceso:"isla", total:16304806, estado:"descartable", web:"https://www.booking.com/hotel/co/isla-grande-islas-del-rosario.html"},
  {id:"marazao", nombre:"Marazao Beach", zona:"Ubicación por confirmar", acceso:"isla", total:16919203, estado:"descartable", web:""},
  {id:"tintipan", nombre:"Hotel Tintipán", zona:"Isla Tintipán · San Bernardo", acceso:"isla", total:19160000, estado:"descartable", web:"https://hoteltintipan.com/"},
  {id:"river", nombre:"Hotel River City", zona:"Montería · Córdoba", acceso:"ruta", total:2600000, estado:"complemento", web:"https://www.google.com/maps/search/?api=1&query=Hotel+River+City+Carrera+5+34-68+Monteria"}
];

const TAREAS = [
  {id:"p2", prio:"alta", t:"Laguna Beach: confirmar la quinta noche", q:"Angela"},
  {id:"p3", prio:"media", t:"Laguna Beach: pedir la mezcla de suites por escrito", q:"Angela"},
  {id:"p4", prio:"alta", t:"Averiguar el nombre real del hotel de El Porvenir", q:"Jorge"},
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

function webOk_(u) {
  const s = String(u == null ? "" : u).trim();
  if (!/^https?:\/\//i.test(s)) return "";
  return s.substring(0, 250);
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
  const nRaw = Number(data.noches);
  const noches = isFinite(nRaw) && nRaw >= 1 && nRaw <= 14 ? Math.round(nRaw) : 4;
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
    noches: noches,
    cotizado: textoCorto_(data.cotizado, 500),
    detalleCosto: textoCorto_(data.detalleCosto, 500),
    comida: textoCorto_(data.comida, 160),
    estado: textoCorto_(data.estado, 40),
    cabe: cabe,
    fuente: textoCorto_(data.fuente, 120),
    mapsq: textoCorto_(data.mapsq, 160),
    wa: textoCorto_(data.wa, 200),
    web: webOk_(data.web),
    si: si,
    no: no,
    extra: extra
  };
}

function escribirPlanes_(planes) {
  const ss = libro_();
  const sh = hoja_(ss, "Planes");
  const headers = ["id", "nombre", "zona", "tipo", "acceso", "horas", "noches", "total", "cotizado", "detalle_costo", "comida", "estado", "caben_10", "fuente", "maps", "whatsapp", "web", "a_favor", "alertas", "letra_menuda"];
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
      o.noches != null ? o.noches : 4,
      o.total != null ? o.total : (cat.total != null ? cat.total : ""),
      o.cotizado || "",
      o.detalleCosto || "",
      o.comida || "",
      o.estado || cat.estado || "",
      cabe,
      o.fuente || "",
      o.mapsq || "",
      o.wa || "",
      o.web || cat.web || "",
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

  sh.getRange(1, 1).setValue("VOTOS — base de datos de la votación familiar");
  sh.getRange(2, 1).setValue("Detalle por votante y plan, y debajo la sumatoria de estrellas. El tablero reescribe esta hoja.");
  sh.getRange(3, 1).clearContent();

  const lastRow = sh.getLastRow();
  const lastCol = Math.max(sh.getLastColumn(), 7);
  if (lastRow >= 4) {
    sh.getRange(4, 1, lastRow - 3, lastCol).clearContent();
  }

  sh.getRange(4, 1, 1, 5).setValues([["Fecha y hora", "Votante", "ID opción", "Opción", "Estrellas (1-5)"]]);

  const filas = [];
  reales.forEach(function (x) {
    const v = x.v;
    const t = ts_(v.ts);
    CATALOGO.forEach(function (o) {
      const e = (v.estrellas && v.estrellas[o.id]) || 0;
      if (!e) return;
      filas.push([t, v.nombre || x.id, o.id, o.nombre, e]);
    });
  });

  let r = 5;
  if (filas.length) {
    sh.getRange(r, 1, filas.length, 5).setValues(filas);
    r += filas.length;
  }

  r += 1;
  sh.getRange(r, 1).setValue("SUMATORIA DE ESTRELLAS");
  r += 1;
  sh.getRange(r, 1, 1, 4).setValues([["Opción", "Votos", "Suma de estrellas", "Promedio"]]);
  r += 1;

  const resumen = [];
  let sumaTotal = 0;
  let votosTotal = 0;
  CATALOGO.forEach(function (o) {
    let suma = 0;
    let n = 0;
    reales.forEach(function (x) {
      const e = (x.v.estrellas && x.v.estrellas[o.id]) || 0;
      if (!e) return;
      suma += e;
      n += 1;
    });
    sumaTotal += suma;
    votosTotal += n;
    resumen.push([o.nombre, n, suma, n ? Math.round((suma / n) * 10) / 10 : 0]);
  });
  resumen.sort(function (a, b) { return b[2] - a[2] || b[3] - a[3]; });
  if (resumen.length) {
    sh.getRange(r, 1, resumen.length, 4).setValues(resumen);
    r += resumen.length;
  }
  sh.getRange(r, 1, 1, 4).setValues([["Total", votosTotal, sumaTotal, votosTotal ? Math.round((sumaTotal / votosTotal) * 10) / 10 : 0]]);
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

function params_(e) {
  const p = {};
  const src = (e && e.parameter) || {};
  Object.keys(src).forEach(function (k) {
    p[k] = src[k];
  });
  if (e && e.postData && e.postData.contents) {
    try {
      const j = JSON.parse(e.postData.contents);
      if (j && typeof j === "object") {
        Object.keys(j).forEach(function (k) {
          if (k === "data" && j.data != null && typeof j.data !== "string") {
            p.data = JSON.stringify(j.data);
          } else if (p[k] == null || p[k] === "") {
            p[k] = j[k];
          }
        });
      }
    } catch (err) {}
  }
  return p;
}

function doGet(e) {
  return despachar_(params_(e), e);
}

function doPost(e) {
  return despachar_(params_(e), e);
}

function despachar_(p, e) {
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
