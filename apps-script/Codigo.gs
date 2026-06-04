/**
 * Bautizo de José Antonio — Backend de confirmaciones (RSVP)
 * Guarda las confirmaciones en una hoja de Google y permite leerlas.
 *
 * Cómo instalarlo (una sola vez):
 *  1. Entra a https://sheets.google.com y crea una hoja nueva
 *     (puedes llamarla "Confirmaciones Bautizo").
 *  2. En el menú: Extensiones → Apps Script.
 *  3. Borra lo que venga y pega TODO este archivo. Guarda (icono de disquete).
 *  4. Arriba a la derecha: Implementar → Nueva implementación.
 *     - Tipo: "Aplicación web".
 *     - Ejecutar como: "Yo".
 *     - Quién tiene acceso: "Cualquier usuario".
 *     - Clic en "Implementar" y autoriza los permisos.
 *  5. Copia la "URL de la aplicación web" (termina en /exec) y envíamela.
 */

const SHEET_NAME = 'Invitados';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    sheet.appendRow([
      new Date(),
      String(data.nombre || '').slice(0, 80),
      Number(data.personas) || 1,
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
    return { fecha: r[0], nombre: r[1], personas: r[2], mensaje: r[3] };
  });
  const payload = JSON.stringify({ ok: true, invitados: rows });

  const cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService
      .createTextOutput(cb + '(' + payload + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Fecha', 'Nombre', 'Personas', 'Mensaje']);
  }
  return sheet;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
