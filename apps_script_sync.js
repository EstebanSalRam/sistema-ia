// ============================================================
// GOOGLE APPS SCRIPT — Servidor de sincronización
// para los 3 artefactos del Sistema IA de Gerencia
// ============================================================
// INSTRUCCIONES DE INSTALACIÓN:
// 1. Abre Google Sheets → crea un Sheet nuevo en blanco
// 2. Menú: Extensiones → Apps Script
// 3. Borra todo el código que aparece y pega este archivo completo
// 4. Haz clic en "Guardar" (ícono de disco)
// 5. Haz clic en "Implementar" → "Nueva implementación"
// 6. Tipo: "Aplicación web"
// 7. Ejecutar como: "Yo (tu cuenta Google)"
// 8. Quién tiene acceso: "Cualquier usuario" (necesario para iOS)
// 9. Haz clic en "Implementar" → copia la URL que aparece
// 10. Pega esa URL en los 3 artefactos cuando te lo pidan
// ============================================================

const SHEET_NAME_BRIEFING   = 'briefing';
const SHEET_NAME_SEMANAL    = 'estado_semanal';
const SHEET_NAME_COMANDO    = 'centro_comando';

function doGet(e) {
  try {
    const artefacto = e.parameter.artefacto;
    const clave     = e.parameter.clave;
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const sheet     = getOrCreateSheet(ss, artefacto);
    const data      = leerDatos(sheet, clave);
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
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const sheet     = getOrCreateSheet(ss, artefacto);
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
  const datos  = sheet.getDataRange().getValues();
  const valorStr = typeof valor === 'string' ? valor : JSON.stringify(valor);
  const ahora  = new Date().toISOString();
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
// FIN DEL SCRIPT
// ============================================================
