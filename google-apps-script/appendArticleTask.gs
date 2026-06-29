/**
 * Logs a generated draft to the Drafts Log sheet in this spreadsheet.
 * If the sheet doesn't exist, it automatically creates it with headers.
 *
 * Columns: Date | Draft Title | Google Doc Link | Status | Form Row
 * "Form Row" is hidden — it stores the row number in the Form Responses sheet
 * so failed drafts can be reprocessed automatically.
 *
 * @param {string} title    - title of the draft (or failure label)
 * @param {string} docUrl   - link to the generated Google Doc (empty string on failure)
 * @param {string} [status] - '✅ Done' (default) or '⚠️ FAILED'
 * @param {number} [formRowNum] - row index in the Form Responses sheet (for retry)
 * @return {number} row index of the logged entry
 */
function logDraftToSheet(title, docUrl, status, formRowNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = Config.DRAFT_LOG_SHEET_NAME || 'Drafts Log';
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    _initLogHeaders(sheet);
  }

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var rowStatus = status || '✅ Done';
  var rowFormRef = formRowNum || '';
  sheet.appendRow([today, title, docUrl, rowStatus, rowFormRef]);
  return sheet.getLastRow();
}

/**
 * Applies bold headers and column widths to a freshly created log sheet.
 * Hides column E (Form Row) — it's internal plumbing, not for editors.
 * @param {Sheet} sheet
 */
function _initLogHeaders(sheet) {
  sheet.appendRow(['Date', 'Draft Title', 'Google Doc Link', 'Status', 'Form Row']);
  sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 280);
  sheet.setColumnWidth(3, 380);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 80);
  sheet.hideColumns(5); // Form Row is internal — keep the UI clean
}

/**
 * Developer utility to test logging to the local spreadsheet.
 */
function testLogDraftToSheet() {
  var row = logDraftToSheet('TEST – Universal Newsletter Draft', 'https://docs.google.com/document/d/example/edit');
  Logger.log('Logged test draft at row ' + row + ' in the "Drafts Log" sheet.');
}
