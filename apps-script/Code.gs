// Votación camisetas Salvia Blanca 2026
// Pegar en la hoja de Google: Extensiones > Apps Script. Implementar > Nueva implementación > Aplicación web
// Ejecutar como: Yo. Quién tiene acceso: Cualquier persona.
var LABELS = {
  o1: 'Opción 1 · Pulpo en plumilla',
  o2: 'Opción 2 · Estampilla de langosta',
  o3: 'Opción 3 · Sellos verde azulado'
};

function hoja_() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (s.getLastRow() === 0) s.appendRow(['Fecha y hora', 'Nombre', '1er lugar', '2do lugar', '3er lugar']);
  return s;
}

function idDe_(label) {
  for (var k in LABELS) if (LABELS[k] === label) return k;
  return '';
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function notas_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('Notas');
  if (!s) { s = ss.insertSheet('Notas'); s.appendRow(['Fecha y hora', 'Nombre', 'Nota']); }
  return s;
}

// Devuelve votos (nombre, hora y lugares) y notas (nombre, hora y texto)
function doGet() {
  var rows = hoja_().getDataRange().getValues().slice(1);
  var votes = rows.map(function (r) {
    return { name: String(r[1]), at: new Date(r[0]).getTime(), first: idDe_(r[2]), second: idDe_(r[3]), third: idDe_(r[4]) };
  }).filter(function (v) { return v.first && v.second && v.third; });
  var nrows = notas_().getDataRange().getValues().slice(1).slice(-100);
  var notes = nrows.map(function (r) {
    return { name: String(r[1]), text: String(r[2]), at: new Date(r[0]).getTime() };
  });
  return json_({ votes: votes, notes: notes });
}

// Un voto por nombre: si el nombre ya votó, se actualiza su fila
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var d = JSON.parse(e.postData.contents);
    var name = String(d.name || '').trim().slice(0, 40);
    if (d.type === 'note') {
      var texto = String(d.text || '').trim().slice(0, 300);
      if (!name || !texto) return json_({ ok: false });
      var n = notas_();
      n.appendRow([new Date(), name, "'" + texto]);
      n.getRange(n.getLastRow(), 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
      return json_({ ok: true });
    }
    var p = [d.first, d.second, d.third];
    var valido = name && p.every(function (x) { return LABELS[x]; }) && p[0] !== p[1] && p[0] !== p[2] && p[1] !== p[2];
    if (!valido) return json_({ ok: false });
    var s = hoja_();
    var rows = s.getDataRange().getValues();
    var fila = -1;
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][1]).trim().toLowerCase() === name.toLowerCase()) { fila = i + 1; break; }
    }
    var vals = [new Date(), name, LABELS[p[0]], LABELS[p[1]], LABELS[p[2]]];
    if (fila > 0) s.getRange(fila, 1, 1, 5).setValues([vals]);
    else { s.appendRow(vals); fila = s.getLastRow(); }
    s.getRange(fila, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false });
  } finally {
    lock.releaseLock();
  }
}
