// === Code.gs ===
/**
 * Router for GET requests and HTML includes.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate().setTitle('WebApp');
}

/**
 * Includes an HTML partial file.
 * @param {string} filename
 * @return {string}
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Returns login page (for menu). */
function showLogin() {
  return HtmlService.createTemplateFromFile('index').evaluate();
}

/** Menu contextual para recargar la app (solo desarrollo) */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('WebApp')
    .addItem('Recargar\u00A0App', 'showLogin')
    .addToUi();
}

globalThis.showLogin = showLogin;
// === End Code.gs ===
