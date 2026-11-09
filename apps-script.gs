// Pega este código en Extensiones > Apps Script de tu hoja de Google.
// Después: Implementar > Nueva implementación > Aplicación web,
// "Ejecutar como: Yo" y "Quién tiene acceso: Cualquier usuario".
//
// Una fila por jugador (por móvil). Se crea al escribir el nombre y se
// actualiza tras cada pregunta. Para empezar una partida nueva, borra
// las filas de la hoja "Puntuaciones" (deja la cabecera).
const SHEET = 'Puntuaciones';
const HEAD = ['ID', 'Nombre', 'Puntos', 'Aciertos', 'Pregunta', 'Actualizado'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET);
  if (!s) s = ss.insertSheet(SHEET);
  if (s.getLastRow() === 0) s.appendRow(HEAD);
  return s;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  const v = sheet_().getDataRange().getValues().slice(1);
  return json_(v.filter(r => r[0] !== '' && r[1] !== '').map(r => ({
    id: String(r[0]), name: String(r[1]), score: Number(r[2]) || 0, hits: Number(r[3]) || 0, q: Number(r[4]) || 0
  })));
}

function doPost(e) {
  let d = {};
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false }); }
  const id = String(d.id || '').replace(/[^a-z0-9]/gi, '').slice(0, 40);
  let name = String(d.name || '').replace(/\s+/g, ' ').trim().slice(0, 24);
  if (!id || name.length < 2) return json_({ ok: false });
  if (/^[=+\-@]/.test(name)) name = "'" + name;
  const score = Math.max(0, Math.min(15000, Math.round(Number(d.score) || 0)));
  const hits = Math.max(0, Math.min(10, Math.round(Number(d.hits) || 0)));
  const q = Math.max(0, Math.min(10, Math.round(Number(d.q) || 0)));

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const s = sheet_();
    const n = s.getLastRow() - 1;
    const ids = n > 0 ? s.getRange(2, 1, n, 1).getValues().map(r => String(r[0])) : [];
    const i = ids.indexOf(id);
    const row = [id, name, score, hits, q, new Date()];
    if (i < 0) {
      s.appendRow(row);
    } else {
      const prev = s.getRange(i + 2, 1, 1, 5).getValues()[0];
      // Si ya terminó una partida con más puntos, se conserva la mejor.
      if (Number(prev[4]) >= 10 && Number(prev[2]) > score) {
        s.getRange(i + 2, 2).setValue(name);
      } else {
        s.getRange(i + 2, 1, 1, 6).setValues([row]);
      }
    }
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}
