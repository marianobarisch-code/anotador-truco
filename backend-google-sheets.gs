/**
 * Anotador de Truco — backend con Google Sheets
 * ------------------------------------------------
 * Pegá este código en: tu planilla -> Extensiones -> Apps Script.
 * Pasos:
 *   1) TOKEN ya está puesto (debe coincidir con el del cliente).
 *   2) Ejecutá una vez la función setup()  (crea/ordena las 3 pestañas).
 *   3) Implementar -> Nueva implementación (o editar la existente) -> Aplicación web
 *      "Quién tiene acceso": Cualquier persona. Copiá la URL /exec.
 *
 * Métricas de plata: por persona. Si ganás y el rival quedó en las malas
 * (0–14 puntos) se cobra el monto DOBLE; si llegó a buenas (>=15), el SIMPLE.
 * El cliente manda el "monto" ya resuelto por partida.
 */

const TOKEN = 'truco-barisch-7H2k9';

const T_JUG = 'jugadores';
const T_PAR = 'partidas';
const T_DET = 'partidas_jugadores';

const H_JUG = ['id_jugador','nombre','apellido','jugados','ganados','perdidos','puntos','promedio','fecha_alta'];
const H_PAR = ['id_partida','fecha','torneo','modo','equipo_1','equipo_2','ganador','puntos','apuesta_simple','apuesta_doble','monto'];
const H_DET = ['id_partida','id_jugador','equipo','resultado'];

/* === Crea las pestañas con sus títulos (ejecutar una vez) === */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, T_JUG, H_JUG);
  ensureSheet_(ss, T_PAR, H_PAR);
  ensureSheet_(ss, T_DET, H_DET);
}
function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);   // siempre reescribe encabezados
}

/* === LEER (ranking + historial) === */
function doGet(e) {
  if ((e.parameter.token || '') !== TOKEN) return json_({ ok:false, error:'token' });
  return json_({ ok:true, jugadores: leerTabla_(T_JUG), partidas: leerTabla_(T_PAR) });
}

/* === GUARDAR partido === */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents);
    if ((body.token || '') !== TOKEN) return json_({ ok:false, error:'token' });
    if (body.action === 'guardar_partida') return json_(guardarPartida_(body));
    return json_({ ok:false, error:'accion desconocida' });
  } catch (err) {
    return json_({ ok:false, error:String(err) });
  } finally {
    lock.releaseLock();
  }
}

function guardarPartida_(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setup();
  const shJ = ss.getSheetByName(T_JUG);
  const shP = ss.getSheetByName(T_PAR);
  const shD = ss.getSheetByName(T_DET);

  // 1) Asegurar jugadores (crear los nuevos) y resolver sus IDs
  const eq = { 1: [], 2: [] };
  [1, 2].forEach(function (t) {
    (body.equipos[t] || []).forEach(function (p) {
      eq[t].push(asegurarJugador_(shJ, p));
    });
  });

  // 2) Registrar la partida
  const idPartida = 'P' + shP.getLastRow();            // P1, P2, ...
  shP.appendRow([
    idPartida, new Date(), body.torneo || '', body.modo || '',
    eq[1].join(','), eq[2].join(','), 'Equipo ' + body.ganador, body.puntos || '',
    body.apuesta_simple || '', body.apuesta_doble || '', body.monto || ''
  ]);

  // 3) Detalle por jugador + actualizar métricas (victorias, NO plata)
  [1, 2].forEach(function (t) {
    const gano = (t === Number(body.ganador));
    eq[t].forEach(function (id) {
      shD.appendRow([ idPartida, id, t, gano ? 'Ganó' : 'Perdió' ]);
      actualizarMetricas_(shJ, id, gano);
    });
  });

  return { ok:true, id_partida: idPartida, jugadores: leerTabla_(T_JUG), partidas: leerTabla_(T_PAR) };
}

// Mismo nombre+apellido = misma persona (no duplica jugadores)
function asegurarJugador_(shJ, p) {
  const data = shJ.getDataRange().getValues();
  if (p.id) {
    for (var i = 1; i < data.length; i++) if (String(data[i][0]) === String(p.id)) return data[i][0];
  }
  const nom = norm_(p.nombre), ape = norm_(p.apellido);
  for (var j = 1; j < data.length; j++) {
    if (norm_(data[j][1]) === nom && norm_(data[j][2]) === ape) return data[j][0];
  }
  const id = 'J' + data.length;                        // J1, J2, ...
  shJ.appendRow([ id, p.nombre || '', p.apellido || '', 0, 0, 0, 0, '0%', new Date() ]);
  return id;
}
function norm_(s) { return String(s == null ? '' : s).trim().toLowerCase(); }

function actualizarMetricas_(shJ, id, gano) {
  const data = shJ.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      var jug = Number(data[i][3]) + 1;
      var gan = Number(data[i][4]) + (gano ? 1 : 0);
      var per = Number(data[i][5]) + (gano ? 0 : 1);
      var prom = jug ? Math.round(gan / jug * 100) + '%' : '0%';
      shJ.getRange(i + 1, 4, 1, 5).setValues([[ jug, gan, per, gan, prom ]]);
      return;
    }
  }
}

function leerTabla_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getDataRange().getValues();
  const head = data[0];
  return data.slice(1).map(function (r) {
    var o = {}; head.forEach(function (h, i) { o[h] = r[i]; }); return o;
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* === Borra TODOS los datos (deja solo encabezados). Correr desde el editor. === */
function _reset() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [T_JUG, T_PAR, T_DET].forEach(function (n) {
    var sh = ss.getSheetByName(n); var last = sh.getLastRow();
    if (last > 1) sh.deleteRows(2, last - 1);
  });
}
