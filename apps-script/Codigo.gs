/**
 * Bautizo de José Antonio — Backend de confirmaciones (RSVP)
 * Guarda las confirmaciones en una hoja de Google y permite leerlas.
 *
 * Para ACTUALIZAR cuando ya lo tenías instalado:
 *  1. Abre tu hoja → Extensiones → Apps Script.
 *  2. Borra todo y pega este archivo. Guarda.
 *  3. Implementar → Gestionar implementaciones → ✏️ (editar) →
 *     Versión: "Nueva versión" → Implementar.  (La URL NO cambia.)
 */

const SHEET_NAME = 'Invitados';
const HEADERS = ['Fecha', 'Nombre', 'Asistencia', 'Personas', 'Acompañantes', 'Mensaje'];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    const asistencia = (String(data.asistencia).toLowerCase() === 'no') ? 'No' : 'Sí';
    sheet.appendRow([
      new Date(),
      String(data.nombre || '').slice(0, 80),
      asistencia,
      Number(data.personas) || 0,
      String(data.acompanantes || '').slice(0, 400),
      String(data.mensaje || '').slice(0, 200)
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1).map(function (r) {
    return {
      fecha: r[0],
      nombre: r[1],
      asistencia: r[2] || 'Sí',
      personas: r[3],
      acompanantes: r[4] || '',
      mensaje: r[5] || ''
    };
  });
  const payload = JSON.stringify({ ok: true, invitados: rows });
  const cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + payload + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
