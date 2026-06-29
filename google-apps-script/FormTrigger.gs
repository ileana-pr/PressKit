/**
 * Adds a "📋 PressKit" menu to the spreadsheet toolbar when the sheet is opened.
 * Vibe coders can use this menu instead of the Apps Script function dropdown.
 */
function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('📋 PressKit', [
    { name: '▶ Run Pipeline on Last Response', functionName: 'testWithLastResponse' },
    { name: '🔄 Reprocess Failed Drafts', functionName: 'reprocessFailed' },
    { name: '🧪 Submit Test Form Data', functionName: 'seedTestData' }
  ]);
}

// ---------------------------------------------------------------------------
// TEST RUNNER
// ---------------------------------------------------------------------------

/**
 * Test: run the pipeline using the last filled row in the active sheet.
 * Open the form response sheet, then use the "📋 PressKit" menu or
 * the Apps Script editor (Run → testWithLastResponse).
 */
function testWithLastResponse() {
  Logger.log('testWithLastResponse: starting');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ss.toast('No data rows found. Submit the form first, then run again.');
    return;
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var values = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log('testWithLastResponse: lastRow=' + lastRow + ', building article');

  try {
    var article = buildArticleFromResponse(values, headers);
    if (!article || !article.body) {
      ss.toast('No article content found in that row.');
      return;
    }

    Logger.log('testWithLastResponse: calling Gemini...');
    var generated = generateArticle(article.body);

    // ── GUARD: if AI failed, log failure and stop — do NOT create docs ──
    if (!generated.body) {
      logDraftToSheet('⚠️ AI Failed – pending retry', '', '⚠️ FAILED', lastRow);
      ss.toast('⚠️ Gemini is unavailable right now. This submission has been saved as Failed in the Drafts Log. Open the "📋 PressKit" menu and click "🔄 Reprocess Failed Drafts" to retry when Gemini is back.', 'AI Unavailable', 10);
      return;
    }

    var docTitle = generated.title || Config.DEFAULT_TITLE || 'Article draft';
    var draftBody = generated.body;
    Logger.log('testWithLastResponse: title = "' + docTitle + '"');

    var qaDocUrl = createArticleDoc(docTitle, article.body, ' - Q&A');
    var docUrl = createArticleDoc(docTitle, draftBody, '', article.visualAssets);
    Logger.log('testWithLastResponse: draft=' + docUrl + ' | Q&A=' + qaDocUrl);

    logDraftToSheet(docTitle, docUrl, '✅ Done', lastRow);
    ss.toast('Draft created and logged! Check the Drafts Log and your Drive folder. 🎉');
  } catch (err) {
    Logger.log('testWithLastResponse ERROR: ' + err.message);
    Logger.log(err.stack || err.toString());
    ss.toast('Error: ' + err.message);
  }
}

// ---------------------------------------------------------------------------
// LIVE TRIGGER (fires automatically on every new form submission)
// ---------------------------------------------------------------------------

/**
 * Runs when a new form response is submitted ("On form submit" trigger).
 * If Gemini fails, logs the row as FAILED for reprocessing — no Q&A dump.
 */
function onFormSubmit(e) {
  if (!e || !e.values || e.values.length === 0) {
    Logger.log('onFormSubmit: no event data. Use testWithLastResponse() to test manually.');
    return;
  }

  Logger.log('Form submitted');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var values = e.values;
  var sheet = e.range ? e.range.getSheet() : ss.getActiveSheet();
  var formRowNum = e.range ? e.range.rowStart : '';
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  try {
    var article = buildArticleFromResponse(values, headers);
    if (!article || !article.body) {
      ss.toast('No article content found in submission.');
      return;
    }

    Logger.log('onFormSubmit: calling Gemini...');
    var generated = generateArticle(article.body);

    // ── GUARD: if AI failed, log failure and stop — do NOT create docs ──
    if (!generated.body) {
      logDraftToSheet('⚠️ AI Failed – pending retry', '', '⚠️ FAILED', formRowNum);
      ss.toast('⚠️ Gemini unavailable. Submission saved as Failed — use the PressKit menu to retry.', 'AI Unavailable', 10);
      return;
    }

    var docTitle = generated.title || Config.DEFAULT_TITLE || 'Article draft';
    var draftBody = generated.body;
    Logger.log('onFormSubmit: title = "' + docTitle + '"');

    createArticleDoc(docTitle, article.body, ' - Q&A');
    var docUrl = createArticleDoc(docTitle, draftBody, '', article.visualAssets);
    Logger.log('onFormSubmit: draft=' + docUrl);

    logDraftToSheet(docTitle, docUrl, '✅ Done', formRowNum);
    ss.toast('Draft created and logged! 🎉');
  } catch (err) {
    Logger.log('onFormSubmit ERROR: ' + err.message);
    Logger.log(err.stack || err.toString());
    ss.toast('Error: ' + err.message);
  }
}

// ---------------------------------------------------------------------------
// REPROCESS FAILED DRAFTS
// ---------------------------------------------------------------------------

/**
 * Scans the Drafts Log for rows marked '⚠️ FAILED', retrieves the original
 * form response by row number, and re-runs the full pipeline for each one.
 * On success the log row is updated in place. Call from the PressKit menu.
 */
function reprocessFailed() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(Config.DRAFT_LOG_SHEET_NAME || 'Drafts Log');
  if (!logSheet) { ss.toast('No Drafts Log found. Run setupPipeline first.'); return; }

  var lastRow = logSheet.getLastRow();
  if (lastRow < 2) { ss.toast('The Drafts Log is empty.'); return; }

  // Read all log rows (skip header row 1)
  // Columns: A=Date, B=Title, C=DocLink, D=Status, E=FormRow
  var data = logSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var failedEntries = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][3] === '⚠️ FAILED') {
      failedEntries.push({ logRow: i + 2, formRowNum: data[i][4] });
    }
  }

  if (failedEntries.length === 0) {
    ss.toast('No failed drafts to reprocess — everything looks good! 🎉');
    return;
  }

  // Find the Form Responses sheet
  var formSheet = _getFormResponseSheet(ss);
  if (!formSheet) {
    ss.toast('Could not find the Form Responses sheet. Make sure the form is still linked.');
    return;
  }

  var headers = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];
  var successCount = 0;
  var failCount = 0;

  ss.toast('Reprocessing ' + failedEntries.length + ' failed draft(s)… this may take a minute.');

  for (var k = 0; k < failedEntries.length; k++) {
    var entry = failedEntries[k];
    var formRowNum = entry.formRowNum;
    var logRow = entry.logRow;

    if (!formRowNum || formRowNum < 2) {
      Logger.log('reprocessFailed: skipping log row ' + logRow + ' — no form row number stored.');
      failCount++;
      continue;
    }

    try {
      var values = formSheet.getRange(formRowNum, 1, 1, formSheet.getLastColumn()).getValues()[0];
      var article = buildArticleFromResponse(values, headers);
      if (!article || !article.body) { failCount++; continue; }

      var generated = generateArticle(article.body);
      if (!generated.body) {
        Logger.log('reprocessFailed: Gemini still failed for log row ' + logRow);
        failCount++;
        continue;
      }

      var docTitle = generated.title || Config.DEFAULT_TITLE || 'Article draft';
      var draftBody = generated.body;

      createArticleDoc(docTitle, article.body, ' - Q&A');
      var docUrl = createArticleDoc(docTitle, draftBody, '', article.visualAssets);

      // Update the log row in place
      logSheet.getRange(logRow, 2).setValue(docTitle);
      logSheet.getRange(logRow, 3).setValue(docUrl);
      logSheet.getRange(logRow, 4).setValue('✅ Done');

      successCount++;
      Logger.log('reprocessFailed: success for log row ' + logRow + ' → ' + docUrl);
    } catch (err) {
      Logger.log('reprocessFailed error for log row ' + logRow + ': ' + err.message);
      failCount++;
    }
  }

  var msg = successCount + ' draft(s) reprocessed successfully!';
  if (failCount > 0) msg += ' ' + failCount + ' still failed — try again later.';
  ss.toast(msg, 'Reprocess Complete');
}

/**
 * Returns the first sheet whose name contains "Form Responses", or null.
 * @param {Spreadsheet} ss
 * @return {Sheet|null}
 */
function _getFormResponseSheet(ss) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().indexOf('Form Responses') !== -1) return sheets[i];
  }
  return null;
}
