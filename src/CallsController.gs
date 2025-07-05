// === CallsController.gs ===
/**
 * Controlador para operaciones sobre las llamadas.
 */
const CallsController = (() => {
  /** Obtiene llamadas para un monitor */
  function getCallsForMonitor(idMonitor) {
    const ss = SpreadsheetApp.openById(SHEET_CALLS_ID);
    const sheet = ss.getSheetByName(CALLS_SHEET);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const idx = headers.indexOf('ID_Monitor');
    return data
      .filter(r => String(r[idx]) === idMonitor)
      .map(r => headers.reduce((o, h, i) => {
        o[h] = r[i];
        return o;
      }, {}));
  }

  /** Guarda una llamada */
  function saveCall(record) {
    const ss = SpreadsheetApp.openById(SHEET_CALLS_ID);
    const sheet = ss.getSheetByName(CALLS_SHEET);
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) return false;
    try {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const row = headers.map(h => record[h] || '');
      const next = sheet.getLastRow() + 1;
      sheet.getRange(next, 1, 1, row.length).setValues([row]);
      return true;
    } finally {
      lock.releaseLock();
    }
  }

  /** Resumen de llamadas por monitor */
  function getSummary(hours) {
    const ss = SpreadsheetApp.openById(SHEET_CALLS_ID);
    const sheet = ss.getSheetByName(CALLS_SHEET);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const idxId = headers.indexOf('ID_Monitor');
    const idxStatus = headers.indexOf('Contest\u00f3');
    const idxTime = headers.indexOf('fechaHoraReal');
    const start = new Date(Date.now() - hours * 3600 * 1000);
    const acc = {};
    data.forEach(r => {
      const t = new Date(r[idxTime]);
      if (t >= start) {
        const id = r[idxId];
        if (!acc[id]) acc[id] = { llamadas: 0, efectivas: 0, noContesto: 0 };
        acc[id].llamadas++;
        if (r[idxStatus] === 'Contest\u00f3') acc[id].efectivas++; else acc[id].noContesto++;
      }
    });
    return Object.keys(acc).map(k => ({
      Monitor: k,
      llamadas: acc[k].llamadas,
      efectivas: acc[k].efectivas,
      noContesto: acc[k].noContesto,
      tiempo: 0
    }));
  }

  return { getCallsForMonitor, saveCall, getSummary };
})();

globalThis.CallsController = CallsController;
// === End CallsController.gs ===
