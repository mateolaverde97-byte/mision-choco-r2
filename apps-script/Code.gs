/**
 * Misión Chocó · Ronda 2 · Propuestas conjuntas
 * Backend en Google Apps Script para la web de GitHub Pages.
 * Guarda una fila por ficha (pestaña "fichas") y un historial de cada cambio (pestaña "historial").
 *
 * Pasos: ver README.md (crear hoja → Extensiones → Apps Script → pegar este código →
 * ejecutar setup() una vez → Implementar como aplicación web con acceso "Cualquier usuario").
 */

const SHEET_FICHAS = 'fichas';
const SHEET_HIST = 'historial';
const HEAD = ['id', 'bloque', 'actividad', 'equipos', 'estado', 'version', 'actualizado', 'editor',
              'costo_total', 'enviado_por', 'enviado_en', 'data_json'];
const HEAD_HIST = ['fecha', 'id', 'version', 'accion', 'editor', 'costo_total', 'data_json'];

function setup() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_FICHAS) || ss.insertSheet(SHEET_FICHAS);
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
  sh.setFrozenRows(1);
  let hi = ss.getSheetByName(SHEET_HIST) || ss.insertSheet(SHEET_HIST);
  hi.getRange(1, 1, 1, HEAD_HIST.length).setValues([HEAD_HIST]).setFontWeight('bold');
  hi.setFrozenRows(1);
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_FICHAS);
  if (!sh) throw new Error('Falta la pestaña "fichas": ejecuta setup() una vez.');
  return sh;
}

function rowToRec_(r) {
  let data = null;
  try { data = r[11] ? JSON.parse(r[11]) : null; } catch (e) { data = null; }
  const t = r[6] instanceof Date ? r[6].getTime() : Number(r[6]) || null;
  return { id: String(r[0]), estado: String(r[4] || 'borrador'), version: Number(r[5]) || 0, actualizado: t,
           editor: String(r[7] || ''), costo: Number(r[8]) || 0, data: data };
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return -1;
}

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'list';
    const sh = sheet_();
    if (action === 'list') {
      const last = sh.getLastRow();
      const rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, HEAD.length).getValues();
      return out_({ ok: true, items: rows.filter(r => r[0]).map(rowToRec_) });
    }
    if (action === 'get') {
      const row = findRow_(sh, String(e.parameter.id || ''));
      if (row < 0) return out_({ ok: true, rec: null });
      return out_({ ok: true, rec: rowToRec_(sh.getRange(row, 1, 1, HEAD.length).getValues()[0]) });
    }
    return out_({ ok: true, ping: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action !== 'save' && action !== 'send') return out_({ ok: false, error: 'acción no válida' });
    const id = String(body.id || '');
    if (!id) return out_({ ok: false, error: 'falta id' });
    const sh = sheet_();
    const row = findRow_(sh, id);
    let cur = null;
    if (row > 0) cur = rowToRec_(sh.getRange(row, 1, 1, HEAD.length).getValues()[0]);

    // Control de versiones: si otra persona guardó después de que yo cargué la ficha, aviso.
    if (cur && !body.force && Number(body.base) < cur.version) {
      return out_({ ok: false, conflict: true, rec: cur });
    }

    const now = new Date();
    const version = (cur ? cur.version : 0) + 1;
    let estado = 'borrador';
    if (action === 'send') estado = 'enviada';
    else if (cur && (cur.estado === 'enviada' || cur.estado === 'modificada')) estado = 'modificada';

    const meta = body.meta || {};
    const prev = row > 0 ? sh.getRange(row, 1, 1, HEAD.length).getValues()[0] : null;
    const values = [id, meta.bloque || '', meta.actividad || '', meta.equipos || '', estado, version, now,
                    body.editor || '', Number(body.costo) || 0,
                    action === 'send' ? (body.editor || '') : (prev ? prev[9] : ''),
                    action === 'send' ? now : (prev ? prev[10] : ''),
                    JSON.stringify(body.data || {})];
    if (row > 0) sh.getRange(row, 1, 1, HEAD.length).setValues([values]);
    else sh.appendRow(values);

    const hi = SpreadsheetApp.getActive().getSheetByName(SHEET_HIST);
    if (hi) hi.appendRow([now, id, version, action, body.editor || '', Number(body.costo) || 0, JSON.stringify(body.data || {})]);

    return out_({ ok: true, rec: { id: id, estado: estado, version: version, actualizado: now.getTime(),
                                   editor: body.editor || '', costo: Number(body.costo) || 0, data: body.data || {} } });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
