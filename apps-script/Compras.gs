/**
 * Misión Chocó 2026 · Compras de materiales
 * Backend en Google Apps Script para la página /compras/ de GitHub Pages.
 *
 * Guarda una fila por material (pestaña "items"), un historial de cambios (pestaña "historial")
 * y los recibos en la carpeta de Drive "Recibos · Compras Chocó 2026".
 *
 * Pasos (ver README.md): hoja nueva → Extensiones → Apps Script → pegar este código →
 * ejecutar setup() una vez → Implementar como aplicación web (Ejecutar como: Yo · Acceso: Cualquier usuario).
 */

const SH_ITEMS = 'items';
const SH_HIST = 'historial';
const FOLDER_NAME = 'Recibos · Compras Chocó 2026';
const HEAD = ['id', 'equipo', 'bloque', 'actividad', 'material', 'cantidad', 'comprado', 'valor_pagado',
              'quien_compro', 'notas', 'recibos_json', 'actualizado', 'editor', 'version'];
const FIELDS = { cantidad: 5, comprado: 6, valor: 7, quien: 8, notas: 9 };
const HEAD_HIST = ['fecha', 'id', 'accion', 'editor', 'cambios_json'];
const MAX_BYTES = 8 * 1024 * 1024;

function setup() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SH_ITEMS) || ss.insertSheet(SH_ITEMS);
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
  sh.setFrozenRows(1);
  const hi = ss.getSheetByName(SH_HIST) || ss.insertSheet(SH_HIST);
  hi.getRange(1, 1, 1, HEAD_HIST.length).setValues([HEAD_HIST]).setFontWeight('bold');
  hi.setFrozenRows(1);
  folder_();
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH_ITEMS);
  if (!sh) throw new Error('Falta la pestaña "items": ejecuta setup() una vez.');
  return sh;
}

function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function rowToRec_(r) {
  let rec = [];
  try { rec = r[10] ? JSON.parse(r[10]) : []; } catch (e) { rec = []; }
  const t = r[11] instanceof Date ? r[11].getTime() : Number(r[11]) || null;
  return {
    id: String(r[0]),
    cantidad: r[5] === '' || r[5] === null ? null : Number(r[5]),
    comprado: r[6] === true || String(r[6]).toUpperCase() === 'TRUE' || String(r[6]) === 'Sí',
    valor: r[7] === '' || r[7] === null ? null : Number(r[7]),
    quien: String(r[8] || ''),
    notas: String(r[9] || ''),
    recibos: rec,
    actualizado: t,
    editor: String(r[12] || ''),
    version: Number(r[13]) || 0
  };
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return -1;
}

function rowFor_(sh, id, meta) {
  let row = findRow_(sh, id);
  if (row > 0) return row;
  const m = meta || {};
  sh.appendRow([id, m.equipo || '', m.bloque || '', m.actividad || '', m.material || '', '', false, '', '', '', '[]', new Date(), '', 0]);
  return sh.getLastRow();
}

function doGet(e) {
  try {
    const sh = sheet_();
    const last = sh.getLastRow();
    const rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, HEAD.length).getValues();
    return out_({ ok: true, now: Date.now(), items: rows.filter(r => r[0]).map(rowToRec_) });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const id = String(body.id || '');
    if (!id) return out_({ ok: false, error: 'falta id' });

    // Subir el archivo antes de tomar el bloqueo (puede tardar).
    let file = null;
    if (action === 'upload') {
      const bytes = Utilities.base64Decode(String(body.data || ''));
      if (!bytes.length) return out_({ ok: false, error: 'archivo vacío' });
      if (bytes.length > MAX_BYTES) return out_({ ok: false, error: 'archivo muy grande (máx. 8 MB)' });
      const mime = String(body.mime || 'application/octet-stream');
      if (!/^image\/|^application\/pdf$/.test(mime)) return out_({ ok: false, error: 'solo fotos o PDF' });
      const name = (String(body.name || 'recibo')).replace(/[^\w.\- áéíóúñÁÉÍÓÚÑ]/g, '_').slice(0, 80);
      const stamp = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyyMMdd-HHmm');
      file = folder_().createFile(Utilities.newBlob(bytes, mime, id + ' · ' + stamp + ' · ' + name));
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }

    lock.waitLock(20000);
    const sh = sheet_();
    const row = rowFor_(sh, id, body.meta);
    const cur = sh.getRange(row, 1, 1, HEAD.length).getValues()[0];
    const now = new Date();
    const editor = String(body.editor || '');
    let changes = {};

    if (action === 'save') {
      const p = body.patch || {};
      Object.keys(FIELDS).forEach(k => {
        if (!(k in p)) return;
        let v = p[k];
        if (k === 'comprado') v = !!v;
        else if (k === 'cantidad' || k === 'valor') v = (v === null || v === '') ? '' : Number(v);
        else v = String(v).slice(0, 2000);
        cur[FIELDS[k]] = v;
        changes[k] = v;
      });
    } else if (action === 'upload') {
      let rec = [];
      try { rec = cur[10] ? JSON.parse(cur[10]) : []; } catch (x) { rec = []; }
      const r = { fid: file.getId(), url: file.getUrl(), name: String(body.name || 'recibo'), at: now.getTime(), by: editor };
      rec.push(r);
      cur[10] = JSON.stringify(rec);
      changes = { recibo: r.name };
    } else if (action === 'unlink') {
      let rec = [];
      try { rec = cur[10] ? JSON.parse(cur[10]) : []; } catch (x) { rec = []; }
      const fid = String(body.fid || '');
      rec = rec.filter(r => r.fid !== fid);
      cur[10] = JSON.stringify(rec);
      changes = { quitar_recibo: fid };
      try { DriveApp.getFileById(fid).setTrashed(true); } catch (x) {}
    } else {
      return out_({ ok: false, error: 'acción no válida' });
    }

    const m = body.meta || {};
    if (m.equipo) cur[1] = m.equipo;
    if (m.bloque) cur[2] = m.bloque;
    if (m.actividad) cur[3] = m.actividad;
    if (m.material) cur[4] = m.material;
    cur[11] = now;
    cur[12] = editor;
    cur[13] = (Number(cur[13]) || 0) + 1;
    sh.getRange(row, 1, 1, HEAD.length).setValues([cur]);

    const hi = SpreadsheetApp.getActive().getSheetByName(SH_HIST);
    if (hi) hi.appendRow([now, id, action, editor, JSON.stringify(changes)]);

    return out_({ ok: true, rec: rowToRec_(cur) });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}
