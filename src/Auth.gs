// === Auth.gs ===
/**
 * Validates user credentials and returns main interface.
 * @param {string} email
 * @param {string} dni
 * @return {{html:string}|{error:string}}
 */
function login(email, dni) {
  email = Utilities.formatString('%s', email).trim();
  dni = Utilities.formatString('%s', dni).trim();

  const roles = getRolesData();
  const headers = roles[0];
  const idxEmail = headers.indexOf('CORREO');
  const idxDni = headers.indexOf('DNI');
  const idxRol = headers.indexOf('Rol');
  const idxId = headers.indexOf('ID_Monitor');
  const idxName = headers.indexOf('APELLIDOS Y NOMBRES');

  let rowFound = null;
  for (let i = 1; i < roles.length; i++) {
    const row = roles[i];
    if (row[idxEmail] === email && String(row[idxDni]) === dni) {
      rowFound = row;
      break;
    }
  }

  if (!rowFound) {
    logAttempt(email, dni, 'fracaso');
    return { error: 'Matr\u00edcula incorrecta' };
  }

  const role = String(rowFound[idxRol]);
  const idMonitor = String(rowFound[idxId]);
  const nombres = String(rowFound[idxName]);

  logAttempt(email, dni, 'exito');

  const tpl = HtmlService.createTemplateFromFile('main');
  tpl.rol = role;
  tpl.idMonitor = idMonitor;
  tpl.nombres = nombres;
  const html = tpl.evaluate().getContent();
  return { html: html };
}

/** Loads and caches roles data. */
function getRolesData() {
  const cache = CacheService.getScriptCache();
  const key = 'rolesData';
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);
  const ss = SpreadsheetApp.openById(SHEET_ROLES_ID);
  const data = ss.getSheetByName(ROLES_SHEET).getDataRange().getValues();
  cache.put(key, JSON.stringify(data), 300);
  return data;
}

/** Logs login attempts. */
function logAttempt(email, dni, result) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    const ss = SpreadsheetApp.openById(SHEET_ROLES_ID);
    const sheet = ss.getSheetByName(LOGS_SHEET);
    const next = sheet.getLastRow() + 1;
    const values = [[new Date(), email, dni, result]];
    sheet.getRange(next, 1, 1, 4).setValues(values);
  } finally {
    lock.releaseLock();
  }
}

globalThis.login = login;
// === End Auth.gs ===
