// Retiro Marblism: guarda cada respuesta de la encuesta en este Google Sheet.
// Pegar en Extensiones > Apps Script del Sheet y publicar como app web
// (Ejecutar como: yo, Quién tiene acceso: cualquier usuario).

const DATES = ['Jan 31-Feb 7', 'Feb 14-21', 'Feb 21-28', 'Feb 28-Mar 7'];
// Valores de la versión jueves a domingo, por si alguien tiene la página vieja abierta.
const OLD = { 'Feb 4-7': 'Jan 31-Feb 7', 'Feb 18-21': 'Feb 14-21', 'Mar 4-7': 'Feb 28-Mar 7' };

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const name = clean_(d.name);
    const picked = (Array.isArray(d.dates) ? d.dates : []).map(x => OLD[x] || x);
    if (!name || !picked.length) return json_({ ok: false, error: 'missing' });

    const sh = setup_();
    const allergies = clean_(d.allergies) || 'Ninguna';
    sh.appendRow([new Date(), name].concat(DATES.map(x => picked.includes(x) ? 'Sí' : 'No'), [allergies, 'Dom a Dom']));
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Crea las hojas Respuestas y Resumen la primera vez.
function setup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Respuestas');
  if (sh) return sh;

  sh = ss.insertSheet('Respuestas', 0);
  sh.appendRow(['Fecha', 'Nombre'].concat(DATES, ['Alergias', 'Formato']));
  sh.getRange(1, 1, 1, DATES.length + 4).setFontWeight('bold').setBackground('#ffd400');
  sh.setFrozenRows(1);

  const sum = ss.insertSheet('Resumen', 1);
  sum.appendRow(['Fechas', 'Pueden ir']);
  DATES.forEach((x, i) => {
    const col = String.fromCharCode(67 + i); // C, D, E, F
    sum.appendRow([x, '=COUNTIF(Respuestas!' + col + ':' + col + ',"Sí")']);
  });
  sum.appendRow(['Total respuestas', '=COUNTA(Respuestas!B:B)-1']);
  sum.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#ffd400');
  return sh;
}

// Evita que un nombre como "=algo" se ejecute como fórmula.
function clean_(v) {
  const s = String(v || '').trim().slice(0, 200);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
