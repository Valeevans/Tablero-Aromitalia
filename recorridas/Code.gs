/**
 * Aromitalia · Recorridas de Planta (R-P-CA-002.10)
 * API de la app: guarda y lee los datos en esta planilla de Google.
 * Hoja "datos": A = clave · B = valor (JSON) · C = última actualización
 *   rec:<id>    → una recorrida
 *   trat:<id>   → el tratamiento de un desvío (Registro Único de NC)
 *   meta:seed   → marca de que ya se cargó el historial 2026
 * Uso: Implementar > Nueva implementación > Aplicación web · Ejecutar como: yo · Acceso: cualquier persona
 */
const HOJA = 'datos';

function hoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOJA);
  if (!sh) {
    sh = ss.insertSheet(HOJA);
    sh.getRange(1, 1, 1, 3).setValues([['clave', 'valor', 'actualizado']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Lectura: devuelve [[clave, valor], ...]
function doGet() {
  const sh = hoja_();
  const n = sh.getLastRow();
  const rows = n > 1 ? sh.getRange(2, 1, n - 1, 2).getValues().filter(r => r[0] !== '') : [];
  return ContentService.createTextOutput(JSON.stringify(rows)).setMimeType(ContentService.MimeType.JSON);
}

// Escritura: {key, value} guarda o actualiza · {action:'delete', key} elimina
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'bulk') return bulk_(body.items || []);
    const key = String(body.key || '');
    if (!key) return out_({ ok: false, error: 'sin clave' });
    const sh = hoja_();
    const n = sh.getLastRow();
    const keys = n > 1 ? sh.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0])) : [];
    const idx = keys.indexOf(key);
    if (body.action === 'delete') {
      if (idx >= 0) sh.deleteRow(idx + 2);
      return out_({ ok: true });
    }
    const fila = [key, JSON.stringify(body.value), new Date()];
    if (idx >= 0) sh.getRange(idx + 2, 1, 1, 3).setValues([fila]);
    else sh.appendRow(fila);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Carga masiva (historial inicial): una sola escritura en la hoja
function bulk_(items) {
  const sh = hoja_();
  const n = sh.getLastRow();
  const keys = n > 1 ? sh.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0])) : [];
  const nuevas = [];
  items.forEach(it => {
    const fila = [String(it.key), JSON.stringify(it.value), new Date()];
    const idx = keys.indexOf(fila[0]);
    if (idx >= 0) sh.getRange(idx + 2, 1, 1, 3).setValues([fila]); else nuevas.push(fila);
  });
  if (nuevas.length) sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, 3).setValues(nuevas);
  return out_({ ok: true, n: items.length });
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
