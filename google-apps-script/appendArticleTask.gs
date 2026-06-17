/**
 * Logs a generated draft directly to a tab in this same spreadsheet.
 * If the tab doesn't exist, it automatically creates it with headers.
 *
 * @param {string} title - title of the draft
 * @param {string} docUrl - link to the generated Google Doc
 * @return {number} row index of the logged draft
 */
function logDraftToSheet(title, docUrl) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = Config.DRAFT_LOG_SHEET_NAME || 'Drafts Log';
  var sheet = ss.getSheetByName(sheetName);

  // If the log tab doesn't exist, create it with nice headers
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['Date', 'Draft Title', 'Google Doc Link']);
    // Apply basic visual formatting to headers (bold)
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 300);
    sheet.setColumnWidth(3, 400);
  }

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  // Append the new draft row
  sheet.appendRow([today, title, docUrl]);
  return sheet.getLastRow();
}

/**
 * Developer utility to test logging to the local spreadsheet.
 */
function testLogDraftToSheet() {
  var row = logDraftToSheet('TEST – Universal Newsletter Draft', 'https://docs.google.com/document/d/example/edit');
  Logger.log('Logged test draft at row ' + row + ' in the "Drafts Log" tab.');
}
