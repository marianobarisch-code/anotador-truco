/**
 * Anotador de Truco — backend con Google Sheets
 * ------------------------------------------------
 * Pegá este código en: tu planilla -> Extensiones -> Apps Script.
 * Pasos:
 *   1) Cambiá TOKEN por una clave secreta tuya.
 *   2) Ejecutá una vez la función setup()  (crea las 3 pestañas).
 *   3) Implementar -> Nueva implementación -> Aplicación web
 *      "Quién tiene acceso": Cualquier persona. Copiá la URL /exec.
 */

const TOKEN = 'cambia-esta-clave-secreta';   // <-- poné tu clave

const T_JUG = 'jugadores';
const T_PAR = 'partidas';
const T_DET = 'partidas_jugadores';

const H_JUG = ['id_jugador','nombre','apellido','jugados','ganados','perdidos','puntos','promedio','fecha_alta'];
const H_PAR = ['id_partida','fecha','modo','equipo_1','equipo_2','ganador','puntos'];
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
  if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
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
  shP.appendRow([ idPartida, new Date(), body.modo || '', eq[1].join(','), eq[2].join(','), 'Equipo ' + body.ganador, body.puntos || '' ]);

  // 3) Detalle por jugador + actualizar métricas
  [1, 2].forEach(function (t) {
    const gano = (t === Number(body.ganador));
    eq[t].forEach(function (id) {
      shD.appendRow([ idPartida, id, t, gano ? 'Ganó' : 'Perdió' ]);
      actualizarMetricas_(shJ, id, gano);
    });
  });

  return { ok:true, id_partida: idPartida, jugadores: leerTabla_(T_JUG) };
}

function asegurarJugador_(shJ, p) {
  const data = shJ.getDataRange().getValues();
  if (p.id) {
    for (var i = 1; i < data.length; i++) if (String(data[i][0]) === String(p.id)) return data[i][0];
  }
  const id = 'J' + data.length;                        // J1, J2, ...
  shJ.appendRow([ id, p.nombre || '', p.apellido || '', 0, 0, 0, 0, '0%', new Date() ]);
  return id;
}

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
