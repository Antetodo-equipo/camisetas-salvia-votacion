// Votación camisetas Salvia Blanca 2026
// Pegar en la hoja de Google: Extensiones > Apps Script. Implementar > Nueva implementación > Aplicación web
// Ejecutar como: Yo. Quién tiene acceso: Cualquier persona.
// Pestaña 1 "Votos": Fecha y hora, Nombre, 1er, 2do y 3er lugar, Estado (Activo o Cancelado), Grupo.
// Pestaña "Notas": Fecha y hora, Nombre, Nota, Grupo.
// Grupo: antetodo, contemplacion o general (liga sin ?grupo=). Cada grupo ve y cuenta solo lo suyo.
var GRUPOS = ['antetodo', 'contemplacion'];
var LABELS = {
  o1: 'Opción 1 · Pulpo en plumilla',
  o2: 'Opción 2 · Estampilla de langosta',
  o3: 'Opción 3 · Sellos verde azulado'
};
var FMT = 'dd/MM/yyyy HH:mm:ss';

function hoja_() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (s.getLastRow() === 0) s.appendRow(['Fecha y hora', 'Nombre', '1er lugar', '2do lugar', '3er lugar', 'Estado', 'Grupo']);
  else colGrupo_(s, 7);
  return s;
}

function notas_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('Notas');
  if (!s) { s = ss.insertSheet('Notas'); s.appendRow(['Fecha y hora', 'Nombre', 'Nota', 'Grupo']); }
  else colGrupo_(s, 4);
  return s;
}

// Agrega el encabezado "Grupo" a una hoja que ya existía sin esa columna
function colGrupo_(s, col) {
  var h = s.getRange(1, col);
  if (!h.getValue()) {
    h.setValue('Grupo');
    s.getRange(1, col - 1).copyTo(h, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }
}

// Solo grupos conocidos; lo demás cuenta como "general"
function grupo_(g) {
  g = String(g || '').toLowerCase().trim();
  return GRUPOS.indexOf(g) >= 0 ? g : 'general';
}

// Filas sin grupo (anteriores a esta versión) cuentan como "general"
function esGrupo_(celda, g) {
  return grupo_(celda) === g;
}

function idDe_(label) {
  for (var k in LABELS) if (LABELS[k] === label) return k;
  return '';
}

// Mismo nombre aunque cambien mayúsculas, acentos o espacios
function clave_(n) {
  return String(n || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim();
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Marca como Cancelado todo voto activo de ese nombre en ese grupo
function cancelar_(s, key, g) {
  var rows = s.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (clave_(rows[i][1]) === key && esGrupo_(rows[i][6], g) && String(rows[i][5]) !== 'Cancelado') {
      s.getRange(i + 1, 6).setValue('Cancelado');
    }
  }
}

// Devuelve votos activos (nombre, hora y lugares) y notas (nombre, hora y texto) del grupo pedido (?grupo=)
function doGet(e) {
  var g = grupo_(e && e.parameter && e.parameter.grupo);
  if (e && e.parameter && e.parameter.ronda === '2') return doGetV2_(g);
  var rows = hoja_().getDataRange().getValues().slice(1);
  var votes = rows.filter(function (r) { return String(r[5]) !== 'Cancelado' && esGrupo_(r[6], g); }).map(function (r) {
    return { name: String(r[1]), at: new Date(r[0]).getTime(), first: idDe_(r[2]), second: idDe_(r[3]), third: idDe_(r[4]) };
  }).filter(function (v) { return v.first && v.second && v.third; });
  var nrows = notas_().getDataRange().getValues().slice(1).filter(function (r) { return esGrupo_(r[3], g); }).slice(-100);
  var notes = nrows.map(function (r) {
    return { name: String(r[1]), text: String(r[2]), at: new Date(r[0]).getTime() };
  });
  return json_({ votes: votes, notes: notes });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var d = JSON.parse(e.postData.contents);
    var name = String(d.name || '').trim().slice(0, 40);
    if (!name) return json_({ ok: false });
    var g = grupo_(d.grupo);
    if (Number(d.ronda) === 2) return postV2_(d, name, g);

    if (d.type === 'note') {
      var texto = String(d.text || '').trim().slice(0, 300);
      if (!texto) return json_({ ok: false });
      var n = notas_();
      n.appendRow([new Date(), name, "'" + texto, g]);
      n.getRange(n.getLastRow(), 1).setNumberFormat(FMT);
      return json_({ ok: true });
    }

    var s = hoja_();
    if (d.type === 'cancel') {
      cancelar_(s, clave_(name), g);
      return json_({ ok: true });
    }

    var p = [d.first, d.second, d.third];
    var valido = p.every(function (x) { return LABELS[x]; }) && p[0] !== p[1] && p[0] !== p[2] && p[1] !== p[2];
    if (!valido) return json_({ ok: false });
    // El voto anterior de esta persona queda como Cancelado y solo cuenta el nuevo
    cancelar_(s, clave_(name), g);
    s.appendRow([new Date(), name, LABELS[p[0]], LABELS[p[1]], LABELS[p[2]], 'Activo', g]);
    s.getRange(s.getLastRow(), 1).setNumberFormat(FMT);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false });
  } finally {
    lock.releaseLock();
  }
}

// ---------- RONDA 2: dos propuestas, cada persona elige UNA (?ronda=2) ----------
// Pestaña "Votos V2": Fecha y hora, Nombre, Propuesta, Estado (Activo o Cancelado), Grupo.
// Pestaña "Notas V2": Fecha y hora, Nombre, Nota, Grupo.
var PROP = { p1: 'Propuesta 01', p2: 'Propuesta 02' };

function hojaV2_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('Votos V2');
  if (!s) { s = ss.insertSheet('Votos V2'); s.appendRow(['Fecha y hora', 'Nombre', 'Propuesta', 'Estado', 'Grupo']); }
  return s;
}

function notasV2_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('Notas V2');
  if (!s) { s = ss.insertSheet('Notas V2'); s.appendRow(['Fecha y hora', 'Nombre', 'Nota', 'Grupo']); }
  return s;
}

function idPropDe_(label) {
  for (var k in PROP) if (PROP[k] === label) return k;
  return '';
}

function cancelarV2_(s, key, g) {
  var rows = s.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (clave_(rows[i][1]) === key && esGrupo_(rows[i][4], g) && String(rows[i][3]) !== 'Cancelado') {
      s.getRange(i + 1, 4).setValue('Cancelado');
    }
  }
}

function doGetV2_(g) {
  var rows = hojaV2_().getDataRange().getValues().slice(1);
  var votes = rows.filter(function (r) { return String(r[3]) !== 'Cancelado' && esGrupo_(r[4], g); }).map(function (r) {
    return { name: String(r[1]), at: new Date(r[0]).getTime(), choice: idPropDe_(r[2]) };
  }).filter(function (v) { return v.choice; });
  var nrows = notasV2_().getDataRange().getValues().slice(1).filter(function (r) { return esGrupo_(r[3], g); }).slice(-100);
  var notes = nrows.map(function (r) {
    return { name: String(r[1]), text: String(r[2]), at: new Date(r[0]).getTime() };
  });
  return json_({ votes: votes, notes: notes });
}

// Se llama dentro de doPost (ya tiene el candado)
function postV2_(d, name, g) {
  if (d.type === 'note') {
    var texto = String(d.text || '').trim().slice(0, 300);
    if (!texto) return json_({ ok: false });
    var n = notasV2_();
    n.appendRow([new Date(), name, "'" + texto, g]);
    n.getRange(n.getLastRow(), 1).setNumberFormat(FMT);
    return json_({ ok: true });
  }
  var s = hojaV2_();
  if (d.type === 'cancel') {
    cancelarV2_(s, clave_(name), g);
    return json_({ ok: true });
  }
  if (!PROP[d.choice]) return json_({ ok: false });
  // El voto anterior de esta persona queda como Cancelado y solo cuenta el nuevo
  cancelarV2_(s, clave_(name), g);
  s.appendRow([new Date(), name, PROP[d.choice], 'Activo', g]);
  s.getRange(s.getLastRow(), 1).setNumberFormat(FMT);
  return json_({ ok: true });
}
