// Pega este código en Extensiones > Apps Script de tu hoja de Google.
// Después: Implementar > Nueva implementación > Aplicación web,
// "Ejecutar como: Yo" y "Quién tiene acceso: Cualquier usuario".
const SHEET = 'Puntuaciones';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET);
  if (!s) s = ss.insertSheet(SHEET);
  if (s.getLastRow() === 0) s.appendRow(['Fecha', 'Nombre', 'Puntos', 'Aciertos']);
  return s;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  const v = sheet_().getDataRange().getValues().slice(1);
  return json_(v.filter(r => r[1] !== '').map(r => ({ name: String(r[1]), score: Number(r[2]) || 0, hits: Number(r[3]) || 0 })));
}

function doPost(e) {
  let d = {};
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false }); }
  let name = String(d.name || '').replace(/\s+/g, ' ').trim().slice(0, 24);
  if (name.length < 2) return json_({ ok: false });
  if (/^[=+\-@]/.test(name)) name = "'" + name;
  const score = Math.max(0, Math.min(15000, Math.round(Number(d.score) || 0)));
  const hits = Math.max(0, Math.min(10, Math.round(Number(d.hits) || 0)));
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { sheet_().appendRow([new Date(), name, score, hits]); } finally { lock.releaseLock(); }
  return json_({ ok: true });
}
