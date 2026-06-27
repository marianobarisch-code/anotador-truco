/**
 * Anotador de Truco — backend con Google Sheets
 * ------------------------------------------------
 * Pegá este código en: tu planilla -> Extensiones -> Apps Script.
 * Pasos:
 *   1) TOKEN ya está puesto (debe coincidir con el del cliente).
 *   2) Ejecutá una vez la función setup()  (crea/ordena las pestañas).
 *   3) Implementar -> Nueva implementación (o editar la existente) -> Aplicación web
 *      "Quién tiene acceso": Cualquier persona. Copiá la URL /exec.
 *
 * Métricas de plata: por persona. Si ganás y el rival quedó en las malas
 * (0–14 puntos) se cobra el monto DOBLE; si llegó a buenas (>=15), el SIMPLE.
 * El cliente manda el "monto" ya resuelto por partida.
 *
 * Modelo de seguridad: app para grupo de amigos. El acceso lo protege un TOKEN
 * compartido (no hay login). El backend valida la forma de los datos y limita
 * tamaños para que nadie rompa la base con basura, pero quien tenga el token
 * puede escribir — está pensado para gente de confianza.
 */

const TOKEN = 'truco-barisch-7H2k9';

const T_JUG = 'jugadores';
const T_PAR = 'partidas';
const T_DET = 'partidas_jugadores';
const T_JOR = 'jornadas';

const H_JUG = ['id_jugador','nombre','apellido','jugados','ganados','perdidos','puntos','promedio','fecha_alta'];
const H_PAR = ['id_partida','fecha','torneo','modo','equipo_1','equipo_2','ganador','puntos','apuesta_simple','apuesta_doble','monto','jornada'];
const H_DET = ['id_partida','id_jugador','equipo','resultado'];
const H_JOR = ['id_jornada','inicio','fin','estado'];   // estado: abierta | cerrada

/* === Crea las pestañas con sus títulos (ejecutar una vez) === */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, T_JUG, H_JUG);
  ensureSheet_(ss, T_PAR, H_PAR);
  ensureSheet_(ss, T_DET, H_DET);
  ensureSheet_(ss, T_JOR, H_JOR);
}
function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);   // siempre reescribe encabezados
}

/* ============ Helpers de saneo (validación / límites) ============ */
function s_(v, max) { return String(v == null ? '' : v).trim().slice(0, max || 80); }
function num_(v) { var n = Number(v); return isFinite(n) ? n : 0; }
function norm_(s) { return String(s == null ? '' : s).trim().toLowerCase(); }

// Siguiente ID robusto: max(sufijo numérico existente) + 1 (no se rompe si se borran filas).
function nextId_(sh, prefix) {
  var data = sh.getDataRange().getValues();
  var max = 0;
  for (var i = 1; i < data.length; i++) {
    var v = String(data[i][0] || '');
    if (v.indexOf(prefix) === 0) {
      var n = parseInt(v.slice(prefix.length), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  }
  return prefix + (max + 1);
}

/* === LEER (ranking + historial + jornada abierta) === */
function doGet(e) {
  if ((e.parameter.token || '') !== TOKEN) return json_({ ok:false, error:'token' });
  return json_({
    ok: true,
    jugadores: leerTabla_(T_JUG),
    partidas: leerTabla_(T_PAR),
    jornada_abierta: leerJornadaAbierta_()
  });
}

/* === ESCRIBIR (guardar / borrar partida, abrir / cerrar jornada) === */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let body;
    try { body = JSON.parse(e.postData.contents); }
    catch (er) { return json_({ ok:false, error:'json inválido' }); }
    if (!body || typeof body !== 'object') return json_({ ok:false, error:'body inválido' });
    if ((body.token || '') !== TOKEN) return json_({ ok:false, error:'token' });

    switch (body.action) {
      case 'guardar_partida': return json_(guardarPartida_(body));
      case 'borrar_partida':  return json_(borrarPartida_(body));
      case 'abrir_jornada':   return json_(abrirJornada_(body));
      case 'cerrar_jornada':  return json_(cerrarJornada_(body));
      default: return json_({ ok:false, error:'accion desconocida' });
    }
  } catch (err) {
    return json_({ ok:false, error:String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- Partidas ---------------- */
function guardarPartida_(body) {
  // Validación de forma
  var eqIn = body.equipos || {};
  var e1in = eqIn[1] || eqIn['1'] || [];
  var e2in = eqIn[2] || eqIn['2'] || [];
  if (!e1in.length || !e2in.length) return { ok:false, error:'faltan jugadores' };
  if (e1in.length > 6 || e2in.length > 6) return { ok:false, error:'demasiados jugadores' };
  var ganador = Number(body.ganador);
  if (ganador !== 1 && ganador !== 2) return { ok:false, error:'ganador inválido' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setup();
  const shJ = ss.getSheetByName(T_JUG);
  const shP = ss.getSheetByName(T_PAR);
  const shD = ss.getSheetByName(T_DET);

  // 1) Asegurar jugadores (crear los nuevos) y resolver sus IDs
  const eq = { 1: [], 2: [] };
  [1, 2].forEach(function (t) {
    (t === 1 ? e1in : e2in).forEach(function (p) {
      eq[t].push(asegurarJugador_(shJ, p || {}));
    });
  });

  // 2) Registrar la partida (con saneo de strings/números)
  const idPartida = nextId_(shP, 'P');
  shP.appendRow([
    idPartida, new Date(), s_(body.torneo, 24), s_(body.modo, 8),
    eq[1].join(','), eq[2].join(','), 'Equipo ' + ganador, s_(body.puntos, 16),
    num_(body.apuesta_simple), num_(body.apuesta_doble), num_(body.monto), s_(body.jornada, 24)
  ]);

  // 3) Detalle por jugador + actualizar métricas (victorias, NO plata)
  [1, 2].forEach(function (t) {
    const gano = (t === ganador);
    eq[t].forEach(function (id) {
      shD.appendRow([ idPartida, id, t, gano ? 'Ganó' : 'Perdió' ]);
      actualizarMetricas_(shJ, id, gano ? 1 : 0, gano ? 0 : 1);
    });
  });

  return { ok:true, id_partida: idPartida, jugadores: leerTabla_(T_JUG), partidas: leerTabla_(T_PAR) };
}

// Borra una partida: saca su fila, sus filas de detalle y revierte las métricas.
function borrarPartida_(body) {
  var id = s_(body.id_partida, 24);
  if (!id) return { ok:false, error:'falta id_partida' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setup();
  const shJ = ss.getSheetByName(T_JUG);
  const shP = ss.getSheetByName(T_PAR);
  const shD = ss.getSheetByName(T_DET);

  // 1) Revertir métricas y borrar filas de detalle (de abajo hacia arriba)
  var det = shD.getDataRange().getValues();
  var encontrado = false;
  for (var i = det.length - 1; i >= 1; i--) {
    if (String(det[i][0]) === id) {
      var jugId = det[i][1];
      var gano = String(det[i][3]).indexOf('Gan') === 0;   // 'Ganó'
      actualizarMetricas_(shJ, jugId, gano ? -1 : 0, gano ? 0 : -1, -1);
      shD.deleteRow(i + 1);
      encontrado = true;
    }
  }
  // 2) Borrar la fila de la partida
  var par = shP.getDataRange().getValues();
  for (var k = par.length - 1; k >= 1; k--) {
    if (String(par[k][0]) === id) { shP.deleteRow(k + 1); encontrado = true; }
  }
  if (!encontrado) return { ok:false, error:'partida no encontrada' };
  return { ok:true, id_partida:id, jugadores: leerTabla_(T_JUG), partidas: leerTabla_(T_PAR) };
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
  const id = nextId_(shJ, 'J');
  shJ.appendRow([ id, s_(p.nombre, 40), s_(p.apellido, 40), 0, 0, 0, 0, '0%', new Date() ]);
  return id;
}

// dGan/dPer/dJug = deltas (pueden ser negativos al borrar). No baja de 0.
function actualizarMetricas_(shJ, id, dGan, dPer, dJug) {
  if (dJug == null) dJug = 1;
  const data = shJ.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      var jug = Math.max(0, Number(data[i][3]) + dJug);
      var gan = Math.max(0, Number(data[i][4]) + dGan);
      var per = Math.max(0, Number(data[i][5]) + dPer);
      var prom = jug ? Math.round(gan / jug * 100) + '%' : '0%';
      shJ.getRange(i + 1, 4, 1, 5).setValues([[ jug, gan, per, gan, prom ]]);
      return;
    }
  }
}

/* ---------------- Jornadas (sesiones compartidas) ---------------- */
function leerJornadaAbierta_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(T_JOR);
  if (!sh || sh.getLastRow() < 2) return null;
  const data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][3]) === 'abierta') {
      return { id: String(data[i][0]), inicio: data[i][1] };
    }
  }
  return null;
}

// Abre una jornada. Si ya hay una abierta, devuelve esa (no crea otra).
function abrirJornada_(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setup();
  var abierta = leerJornadaAbierta_();
  if (abierta) return { ok:true, jornada: abierta, yaAbierta: true };
  const sh = ss.getSheetByName(T_JOR);
  var inicio = new Date();
  var id = s_(body.id, 24) || String(inicio.getTime());
  sh.appendRow([ id, inicio, '', 'abierta' ]);
  return { ok:true, jornada: { id: id, inicio: inicio } };
}

// Cierra la jornada (la indicada por id, o la que esté abierta).
function cerrarJornada_(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setup();
  const sh = ss.getSheetByName(T_JOR);
  if (sh.getLastRow() < 2) return { ok:true };
  var id = s_(body.id, 24);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var esEsta = id ? String(data[i][0]) === id : String(data[i][3]) === 'abierta';
    if (esEsta && String(data[i][3]) === 'abierta') {
      sh.getRange(i + 1, 3, 1, 2).setValues([[ new Date(), 'cerrada' ]]);
    }
  }
  return { ok:true };
}

/* ---------------- Lectura genérica ---------------- */
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
  [T_JUG, T_PAR, T_DET, T_JOR].forEach(function (n) {
    var sh = ss.getSheetByName(n); if (!sh) return;
    var last = sh.getLastRow();
    if (last > 1) sh.deleteRows(2, last - 1);
  });
}
