/**
 * Test: run the pipeline using the last filled row in the active sheet.
 * Open the form response sheet, run this from the script editor (Run → testWithLastResponse).
 * Creates a Doc and adds a row to the log tab using that row's data.
 */
function testWithLastResponse() {
  Logger.log('testWithLastResponse: starting');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getActiveSpreadsheet().toast('No data rows. Need at least row 1 (headers) and row 2 (one response).');
    return;
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var values = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log('testWithLastResponse: lastRow=' + lastRow + ', building article');

  try {
    var article = buildArticleFromResponse(values, headers);
    if (!article || !article.body) {
      SpreadsheetApp.getActiveSpreadsheet().toast('No article content from that row.');
      return;
    }

    Logger.log('testWithLastResponse: calling Gemini for article and title...');
    var generated = generateArticle(article.body);
    var docTitle = generated.title || Config.DEFAULT_TITLE || 'Article draft';
    var draftBody = generated.body || article.body;
    Logger.log('testWithLastResponse: title = "' + docTitle + '"');

    Logger.log('testWithLastResponse: creating Q&A doc');
    var qaDocUrl = createArticleDoc(docTitle, article.body, ' - Q&A');
    Logger.log('testWithLastResponse: creating draft doc');
    var docUrl = createArticleDoc(docTitle, draftBody, '', article.visualAssets);
    Logger.log('Doc created (draft): ' + docUrl + '; Q&A: ' + qaDocUrl);

    Logger.log('testWithLastResponse: logging draft details to sheet');
    var appendedRow = logDraftToSheet(docTitle, docUrl);
    Logger.log('testWithLastResponse: draft logged at row ' + appendedRow);
    SpreadsheetApp.getActiveSpreadsheet().toast('Draft doc created and logged. Check Drive and the "Drafts Log" tab.');
  } catch (err) {
    Logger.log('testWithLastResponse ERROR: ' + err.message);
    Logger.log(err.stack || err.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast('Error: ' + err.message);
  }
}

/**
 * Runs when a new form response is submitted (install "On form submit" trigger).
 * Builds article from response, creates a Doc, logs draft info to the drafts sheet.
 * For manual testing, use testWithLastResponse() instead.
 */
function onFormSubmit(e) {
  if (!e || !e.values || e.values.length === 0) {
    Logger.log('onFormSubmit: no event data. Use testWithLastResponse() to test with the last row.');
    return;
  }

  Logger.log('Form submitted');
  var values = e.values;
  var sheet = e.range ? e.range.getSheet() : SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  try {
    var article = buildArticleFromResponse(values, headers);
    if (!article || !article.body) {
      SpreadsheetApp.getActiveSpreadsheet().toast('No article content from response.');
      return;
    }

    Logger.log('onFormSubmit: calling Gemini for article and title...');
    var generated = generateArticle(article.body);
    var docTitle = generated.title || Config.DEFAULT_TITLE || 'Article draft';
    var draftBody = generated.body || article.body;
    Logger.log('onFormSubmit: title = "' + docTitle + '"');

    Logger.log('onFormSubmit: creating Q&A doc');
    createArticleDoc(docTitle, article.body, ' - Q&A');
    Logger.log('onFormSubmit: creating draft doc');
    var docUrl = createArticleDoc(docTitle, draftBody, '', article.visualAssets);
    Logger.log('Doc created (draft): ' + docUrl);

    logDraftToSheet(docTitle, docUrl);
    SpreadsheetApp.getActiveSpreadsheet().toast('Draft created and logged to "Drafts Log" sheet.');
  } catch (err) {
    Logger.log('onFormSubmit ERROR: ' + err.message);
    Logger.log(err.stack || err.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast('Error: ' + err.message);
  }
}
