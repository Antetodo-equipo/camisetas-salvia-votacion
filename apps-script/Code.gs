// Votación camisetas Salvia Blanca 2026
// Pegar en la hoja de Google: Extensiones > Apps Script. Implementar > Nueva implementación > Aplicación web
// Ejecutar como: Yo. Quién tiene acceso: Cualquier persona.
// Pestaña 1 "Votos": Fecha y hora, Nombre, 1er, 2do y 3er lugar, Estado (Activo o Cancelado).
// Pestaña "Notas": Fecha y hora, Nombre, Nota.
var LABELS = {
  o1: 'Opción 1 · Pulpo en plumilla',
  o2: 'Opción 2 · Estampilla de langosta',
  o3: 'Opción 3 · Sellos verde azulado'
};
var FMT = 'dd/MM/yyyy HH:mm:ss';

function hoja_() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (s.getLastRow() === 0) s.appendRow(['Fecha y hora', 'Nombre', '1er lugar', '2do lugar', '3er lugar', 'Estado']);
  return s;
}

function notas_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('Notas');
  if (!s) { s = ss.insertSheet('Notas'); s.appendRow(['Fecha y hora', 'Nombre', 'Nota']); }
  return s;
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

// Marca como Cancelado todo voto activo de ese nombre
function cancelar_(s, key) {
  var rows = s.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (clave_(rows[i][1]) === key && String(rows[i][5]) !== 'Cancelado') {
      s.getRange(i + 1, 6).setValue('Cancelado');
    }
  }
}

// Devuelve votos activos (nombre, hora y lugares) y notas (nombre, hora y texto)
function doGet() {
  var rows = hoja_().getDataRange().getValues().slice(1);
  var votes = rows.filter(function (r) { return String(r[5]) !== 'Cancelado'; }).map(function (r) {
    return { name: String(r[1]), at: new Date(r[0]).getTime(), first: idDe_(r[2]), second: idDe_(r[3]), third: idDe_(r[4]) };
  }).filter(function (v) { return v.first && v.second && v.third; });
  var nrows = notas_().getDataRange().getValues().slice(1).slice(-100);
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

    if (d.type === 'note') {
      var texto = String(d.text || '').trim().slice(0, 300);
      if (!texto) return json_({ ok: false });
      var n = notas_();
      n.appendRow([new Date(), name, "'" + texto]);
      n.getRange(n.getLastRow(), 1).setNumberFormat(FMT);
      return json_({ ok: true });
    }

    var s = hoja_();
    if (d.type === 'cancel') {
      cancelar_(s, clave_(name));
      return json_({ ok: true });
    }

    var p = [d.first, d.second, d.third];
    var valido = p.every(function (x) { return LABELS[x]; }) && p[0] !== p[1] && p[0] !== p[2] && p[1] !== p[2];
    if (!valido) return json_({ ok: false });
    // El voto anterior de esta persona queda como Cancelado y solo cuenta el nuevo
    cancelar_(s, clave_(name));
    s.appendRow([new Date(), name, LABELS[p[0]], LABELS[p[1]], LABELS[p[2]], 'Activo']);
    s.getRange(s.getLastRow(), 1).setNumberFormat(FMT);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false });
  } finally {
    lock.releaseLock();
  }
}
