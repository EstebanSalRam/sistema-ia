// ============================================================
// GOOGLE APPS SCRIPT — Servidor de sincronización
// Sistema IA de Gerencia — v2
// ============================================================
// INSTRUCCIONES:
// 1. Abre el Google Sheet vinculado a este script
// 2. Menú: Extensiones → Apps Script
// 3. Reemplaza todo el código con este archivo
// 4. Guarda (Ctrl+S)
// 5. Ejecuta la función "setup" UNA VEZ desde el editor
//    (selecciona "setup" en el desplegable y haz clic en ▶ Ejecutar)
//    Esto guarda el ID de la hoja en las propiedades del script.
// 6. Implementar → Administrar implementaciones → edita la implementación
//    existente → selecciona "Nueva versión" → Implementar
// ============================================================

const SHEET_NAME_BRIEFING = 'briefing';
const SHEET_NAME_SEMANAL  = 'estado_semanal';
const SHEET_NAME_COMANDO  = 'centro_comando';

// Ejecuta esta función UNA VEZ desde el editor de Apps Script
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    Logger.log('ERROR: No se encontró hoja activa. Asegúrate de abrir este script desde una Google Sheet (Extensiones → Apps Script).');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  Logger.log('✓ Setup completado. Sheet ID guardado: ' + ss.getId());
}

function getSpreadsheet() {
  // 1. Intenta con el ID guardado en propiedades (configurado por setup())
  const props = PropertiesService.getScriptProperties();
  const ssId  = props.getProperty('SHEET_ID');
  if (ssId) {
    return SpreadsheetApp.openById(ssId);
  }

  // 2. Fallback: intenta getActiveSpreadsheet (funciona en contexto de editor)
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    props.setProperty('SHEET_ID', active.getId());
    return active;
  }

  // 3. Crea una hoja nueva automáticamente si no hay ninguna configurada
  const newSs = SpreadsheetApp.create('Sistema IA — Sync Data');
  props.setProperty('SHEET_ID', newSs.getId());
  return newSs;
}

function doGet(e) {
  try {
    const artefacto = e.parameter.artefacto;
    const clave     = e.parameter.clave;

    if (!artefacto || artefacto === 'ping') {
      return jsonResponse({ ok: true, msg: 'Sistema IA Sync activo', ts: new Date().toISOString() });
    }

    const ss    = getSpreadsheet();
    const sheet = getOrCreateSheet(ss, artefacto);
    const data  = leerDatos(sheet, clave);
    return jsonResponse({ ok: true, data: data });
  } catch(err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function doPost(e) {
  try {
    const payload   = JSON.parse(e.postData.contents);
    const artefacto = payload.artefacto;
    const clave     = payload.clave;
    const valor     = payload.valor;

    if (artefacto === 'ping') {
      return jsonResponse({ ok: true, msg: 'Sistema IA Sync activo', ts: new Date().toISOString() });
    }

    const ss    = getSpreadsheet();
    const sheet = getOrCreateSheet(ss, artefacto);
    escribirDatos(sheet, clave, valor);
    return jsonResponse({ ok: true, ts: new Date().toISOString() });
  } catch(err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function getOrCreateSheet(ss, nombre) {
  let sheet = ss.getSheetByName(nombre);
  if (!sheet) {
    sheet = ss.insertSheet(nombre);
    sheet.getRange('A1').setValue('clave');
    sheet.getRange('B1').setValue('valor');
    sheet.getRange('C1').setValue('actualizado');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function leerDatos(sheet, clave) {
  const datos = sheet.getDataRange().getValues();
  if (clave) {
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] === clave) {
        try { return JSON.parse(datos[i][1]); }
        catch(e) { return datos[i][1]; }
      }
    }
    return null;
  }
  const resultado = {};
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0]) {
      try { resultado[datos[i][0]] = JSON.parse(datos[i][1]); }
      catch(e) { resultado[datos[i][0]] = datos[i][1]; }
    }
  }
  return resultado;
}

function escribirDatos(sheet, clave, valor) {
  const datos    = sheet.getDataRange().getValues();
  const valorStr = typeof valor === 'string' ? valor : JSON.stringify(valor);
  const ahora    = new Date().toISOString();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === clave) {
      sheet.getRange(i + 1, 2).setValue(valorStr);
      sheet.getRange(i + 1, 3).setValue(ahora);
      return;
    }
  }
  const ultimaFila = sheet.getLastRow() + 1;
  sheet.getRange(ultimaFila, 1).setValue(clave);
  sheet.getRange(ultimaFila, 2).setValue(valorStr);
  sheet.getRange(ultimaFila, 3).setValue(ahora);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
// ============================================================
