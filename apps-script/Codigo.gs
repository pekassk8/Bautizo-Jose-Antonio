/**
 * Bautizo de José Antonio — Backend de confirmaciones (RSVP)
 *
 * Hojas:
 *   - "Invitados"     → los que SÍ asistirán
 *   - "No asistirán"  → los que NO podrán asistir
 * Columnas: Fecha | Invitado | Cantidad | Mensaje para la familia
 *
 * Cada persona se guarda en su propia fila. La "Cantidad" solo aparece
 * en la primera persona del grupo (la cabeza de familia); en los demás
 * acompañantes se deja en blanco para saber que van juntos.
 *
 * Para ACTUALIZAR:
 *  1. Abre tu hoja → Extensiones → Apps Script.
 *  2. Borra todo y pega este archivo. Guarda.
 *  3. Implementar → Gestionar implementaciones → ✏️ (editar) →
 *     Versión: "Nueva versión" → Implementar.  (La URL NO cambia.)
 */

const HOJA_SI = 'Invitados';
const HOJA_NO = 'No asistirán';
const HEADERS = ['Fecha', 'Invitado', 'Cantidad', 'Mensaje para la familia'];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const noAsiste = (String(data.asistencia).toLowerCase() === 'no');
    const fecha = new Date();
    const mensaje = String(data.mensaje || '').slice(0, 200);

    let nombres = Array.isArray(data.nombres)
      ? data.nombres.map(function (n) { return String(n).trim(); }).filter(Boolean)
      : [];
    if (!nombres.length) {
      // Respaldo para versiones antiguas del formulario (nombre + acompañantes)
      nombres = [];
      if (data.nombre) { nombres.push(String(data.nombre).trim()); }
      if (data.acompanantes) {
        String(data.acompanantes).split(',').forEach(function (s) {
          s = s.trim(); if (s) { nombres.push(s); }
        });
      }
      nombres = nombres.filter(Boolean);
      if (!nombres.length) { nombres = ['']; }
    }
    const cantidad = Number(data.personas) || nombres.length || 1;

    const sheet = getSheet_(noAsiste ? HOJA_NO : HOJA_SI);
    if (noAsiste) {
      sheet.appendRow([fecha, nombres[0].slice(0, 80), '', mensaje]);
    } else {
      nombres.forEach(function (nm, idx) {
        sheet.appendRow([fecha, nm.slice(0, 80), idx === 0 ? cantidad : '', mensaje]);
      });
    }
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
    return { fecha: r[0], invitado: r[1], cantidad: (r[2] === '' ? '' : r[2]), mensaje: r[3] || '' };
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
