/**
 * Bautizo de José Antonio — Backend de confirmaciones (RSVP)
 *
 * Guarda en dos hojas:
 *   - "Invitados"     → los que SÍ asistirán
 *   - "No asistirán"  → los que NO podrán asistir
 * Columnas en ambas: Fecha | Invitado | Mensaje para la familia
 *
 * Para ACTUALIZAR si ya lo tenías:
 *  1. Abre tu hoja → Extensiones → Apps Script.
 *  2. Borra todo y pega este archivo. Guarda.
 *  3. Implementar → Gestionar implementaciones → ✏️ (editar) →
 *     Versión: "Nueva versión" → Implementar.  (La URL NO cambia.)
 */

const HOJA_SI = 'Invitados';
const HOJA_NO = 'No asistirán';
const HEADERS = ['Fecha', 'Invitado', 'Mensaje para la familia'];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const noAsiste = (String(data.asistencia).toLowerCase() === 'no');

    let invitado = String(data.nombre || '').slice(0, 80);
    if (data.acompanantes) {
      invitado += ', ' + String(data.acompanantes).slice(0, 400);
    }

    const sheet = getSheet_(noAsiste ? HOJA_NO : HOJA_SI);
    sheet.appendRow([new Date(), invitado, String(data.mensaje || '').slice(0, 200)]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  const payload = JSON.stringify({
    ok: true,
    invitados: leer_(HOJA_SI),
    noAsistiran: leer_(HOJA_NO)
  });
  const cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + payload + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

function leer_(nombreHoja) {
  const sheet = getSheet_(nombreHoja);
  const values = sheet.getDataRange().getValues();
  return values.slice(1).map(function (r) {
    return { fecha: r[0], invitado: r[1], mensaje: r[2] || '' };
  });
}

function getSheet_(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(nombreHoja);
  if (!sheet) { sheet = ss.insertSheet(nombreHoja); }
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
