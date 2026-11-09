// Quiz SL — servidor de la partida (Google Apps Script)
// 1. Pega este código en script.google.com (o Extensiones > Apps Script de una hoja).
// 2. Cambia CLAVE_PROFE por tu propia clave secreta.
// 3. Ejecuta "preparar" una vez y acepta los permisos.
// 4. Implementar > Aplicación web, "Ejecutar como: Yo", "Acceso: Cualquier usuario".
const CLAVE_PROFE = 'CAMBIA-ESTA-CLAVE';

const SHEET = 'Puntuaciones';
const HEAD = ['ID', 'Nombre', 'Puntos', 'Aciertos', 'Pregunta', 'Actualizado'];

function book_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create('Quiz SL - Puntuaciones');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function preparar() {
  const ss = book_();
  sheet_();
  Logger.log('Hoja lista: ' + ss.getUrl());
}

function sheet_() {
  const ss = book_();
  let s = ss.getSheetByName(SHEET);
  if (!s) s = ss.insertSheet(SHEET);
  if (s.getLastRow() === 0) s.appendRow(HEAD);
  return s;
}

function state_() {
  const p = PropertiesService.getScriptProperties();
  return { phase: p.getProperty('PHASE') || 'lobby', game: Number(p.getProperty('GAME') || 1) };
}

function setState_(o) {
  const p = PropertiesService.getScriptProperties();
  if (o.phase) p.setProperty('PHASE', o.phase);
  if (o.game) p.setProperty('GAME', String(o.game));
}

function rows_(s) {
  const n = s.getLastRow() - 1;
  return n > 0 ? s.getRange(2, 1, n, 5).getValues() : [];
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  const st = state_();
  const players = rows_(sheet_()).filter(r => r[0] !== '' && r[1] !== '').map(r => ({
    id: String(r[0]), name: String(r[1]), score: Number(r[2]) || 0, hits: Number(r[3]) || 0, q: Number(r[4]) || 0
  }));
  return json_({ phase: st.phase, game: st.game, players: players });
}

function doPost(e) {
  let d = {};
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'datos' }); }
  const act = String(d.action || '');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const st = state_();
    const s = sheet_();

    // ----- acciones del profesor -----
    if (['check', 'start', 'finish', 'reset'].indexOf(act) >= 0) {
      if (String(d.key || '') !== CLAVE_PROFE) return json_({ ok: false, error: 'clave' });
      if (act === 'start') setState_({ phase: 'playing' });
      if (act === 'finish') setState_({ phase: 'finished' });
      if (act === 'reset') {
        const n = s.getLastRow() - 1;
        if (n > 0) s.deleteRows(2, n);
        setState_({ phase: 'lobby', game: st.game + 1 });
      }
      const ns = state_();
      return json_({ ok: true, phase: ns.phase, game: ns.game });
    }

    // ----- acciones de los jugadores -----
    const id = String(d.id || '').replace(/[^a-z0-9]/gi, '').slice(0, 40);
    if (!id) return json_({ ok: false, error: 'datos' });
    if (Number(d.game) !== st.game) return json_({ ok: false, error: 'partida', game: st.game, phase: st.phase });
    const rows = rows_(s);
    const i = rows.findIndex(r => String(r[0]) === id);

    if (act === 'join') {
      if (st.phase === 'finished') return json_({ ok: false, error: 'terminada' });
      let name = String(d.name || '').replace(/\s+/g, ' ').trim().slice(0, 24);
      if (name.length < 2) return json_({ ok: false, error: 'nombre-corto' });
      if (/^[=+\-@]/.test(name)) name = "'" + name;
      const clash = rows.some(r => String(r[0]) !== id && String(r[1]).toLowerCase() === name.toLowerCase());
      if (clash) return json_({ ok: false, error: 'nombre-usado' });
      if (i < 0) s.appendRow([id, name, 0, 0, 0, new Date()]);
      else s.getRange(i + 2, 2).setValue(name);
      return json_({ ok: true, phase: st.phase, game: st.game });
    }

    if (act === 'progress') {
      if (st.phase !== 'playing') return json_({ ok: false, error: 'fase', phase: st.phase });
      if (i < 0) return json_({ ok: false, error: 'no-unido' });
      if (Number(rows[i][4]) >= 10) return json_({ ok: true });
      const score = Math.max(0, Math.min(15000, Math.round(Number(d.score) || 0)));
      const hits = Math.max(0, Math.min(10, Math.round(Number(d.hits) || 0)));
      const q = Math.max(0, Math.min(10, Math.round(Number(d.q) || 0)));
      s.getRange(i + 2, 3, 1, 4).setValues([[score, hits, q, new Date()]]);
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'accion' });
  } finally {
    lock.releaseLock();
  }
}
