// =============================================================================
// Universal Newsletter Pipeline — Single-File Bundle
// =============================================================================
// HOW TO USE:
//   1. Open your Google Sheet → Extensions ➔ Apps Script
//   2. Select all the code in the default Code.gs file and replace it with this
//   3. Click Save, then select "setupPipeline" in the function dropdown and click Run
//   4. Authorize permissions when prompted
//   5. Check the Execution Log for your Form links — you're done!
//
// REQUIRED: Add your Gemini API key in Project Settings → Script Properties
//   Property name : GEMINI_API_KEY
//   Value         : your key from https://aistudio.google.com/
// =============================================================================


// =============================================================================
// SECTION 1: CONFIG
// Edit DRAFT_FOLDER_ID to save draft Docs to a specific Drive folder.
// Leave it blank ('') to save in your Drive root.
// =============================================================================

/**
 * Master Newsletter Pipeline – General Configuration
 * Universal template for any organization.
 */
var Config = {
  // (Optional) Your organization's name.
  // When set, Gemini will tailor generated test submissions and article tone to your org.
  // Leave blank ('') for generic output that works with any organization.
  ORG_NAME: '',

  // Name given to the spreadsheet when setupPipeline() runs
  SPREADSHEET_NAME: 'Newsletter Pipeline',

  // Name of the sheet/tab in this spreadsheet where generated drafts will be logged
  DRAFT_LOG_SHEET_NAME: 'Drafts Log',

  // Default fallback title if AI generation fails
  DEFAULT_TITLE: 'Untitled Draft',

  // New draft Doc: optional prefix in file name. Date (YYYY-MM-DD) is always prepended.
  DOC_TITLE_PREFIX: '',

  // Folder ID where new draft Google Docs will be created.
  // Leave blank ('') to automatically save in the same folder as this spreadsheet.
  // Set a specific ID to override: drive.google.com/drive/folders/YOUR_ID_HERE
  DRAFT_FOLDER_ID: '',

  // Gemini model for content generation.
  // gemini-2.5-flash-lite is fast, highly capable, and has generous free-tier quotas.
  GEMINI_MODEL: 'gemini-2.5-flash-lite'
};


// =============================================================================
// SECTION 2: ARTICLE BUILDER
// Parses raw form responses into a Q/A block for Gemini.
// =============================================================================

/**
 * Builds Q/A article text from one form response.
 * Title generation is handled by Gemini via generateArticle().
 *
 * @param {Array} values - new row from e.values
 * @param {Array} headers - header row (same length as values)
 * @return {Object} { body, visualAssets }
 */
function buildArticleFromResponse(values, headers) {
  if (!values || !headers || values.length === 0) {
    return { body: '', visualAssets: '' };
  }

  var qaBlocks = [];
  var visualAssets = '';

  for (var i = 0; i < values.length; i++) {
    var question = String(headers[i] || '').trim();
    var value = values[i] !== null && values[i] !== undefined ? String(values[i]).trim() : '';
    if (!value) continue;

    if (question.toLowerCase().indexOf('timestamp') !== -1) continue;

    // Capture the Visual Assets link separately so it can be surfaced in the draft doc
    if (question.toLowerCase().indexOf('visual') !== -1) {
      visualAssets = value;
    }

    qaBlocks.push('Q: ' + (question || '') + '\nA: ' + value);
  }

  return { body: qaBlocks.join('\n\n'), visualAssets: visualAssets };
}

/**
 * Creates a new Google Doc with the given title and body; moves to Config.DRAFT_FOLDER_ID if set.
 * Doc name = date (YYYY-MM-DD) + optional Config.DOC_TITLE_PREFIX + title + optional nameSuffix.
 *
 * @param {string} title - AI-generated title for the article
 * @param {string} body - plain text body; newlines preserved
 * @param {string} [nameSuffix] - optional suffix for file name (e.g. ' - Q&A')
 * @param {string} [visualAssetsUrl] - optional Visual Assets link to append for editors
 * @return {string} doc.getUrl()
 */
function createArticleDoc(title, body, nameSuffix, visualAssetsUrl) {
  var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var suffix = (nameSuffix !== undefined && nameSuffix !== null) ? nameSuffix : '';
  var fullTitle = dateStr + ' - ' + (Config.DOC_TITLE_PREFIX || '') + (title || 'Article draft') + suffix;
  var doc = DocumentApp.create(fullTitle);
  var docBody = doc.getBody();
  var text = (body || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  docBody.setText(text);

  // Append a clearly labeled Visual Assets section so the editor can
  // access the submitted files without opening the original form response.
  if (visualAssetsUrl) {
    docBody.appendParagraph('');
    var divider = docBody.appendParagraph('────────────────────────────────────────');
    divider.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    var label = docBody.appendParagraph('📸 Visual Assets (submitted by contributor)');
    label.setBold(true);
    docBody.appendParagraph(visualAssetsUrl);
  }

  doc.saveAndClose();

  // Move the doc to the right folder:
  // Use Config.DRAFT_FOLDER_ID if explicitly set, otherwise default to the
  // same folder as the spreadsheet so everything stays in one place.
  var targetFolder;
  if (Config.DRAFT_FOLDER_ID) {
    targetFolder = DriveApp.getFolderById(Config.DRAFT_FOLDER_ID);
  } else {
    var ssParents = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents();
    if (ssParents.hasNext()) {
      targetFolder = ssParents.next();
    }
  }
  if (targetFolder) {
    DriveApp.getFileById(doc.getId()).moveTo(targetFolder);
  }

  return doc.getUrl();
}


// =============================================================================
// SECTION 3: GEMINI AI
// Calls the Gemini API to generate the article body and title in one shot.
// =============================================================================

/**
 * Generates both the article body and a matching title from Q/A text using a single Gemini call.
 * The title is written from the finished article context so they are always coherent.
 *
 * @param {string} qaText - Q/A format text from buildArticleFromResponse
 * @return {Object} { title {string}, body {string} } — both fall back to empty string on error
 */
function generateArticle(qaText) {
  var apiKey = getGeminiApiKey();
  if (!apiKey) {
    Logger.log('GeminiArticle: GEMINI_API_KEY not set in Script Properties. Skipping AI.');
    return { title: '', body: '' };
  }

  var model = Config.GEMINI_MODEL || 'gemini-2.5-flash-lite';
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent';
  Logger.log('GeminiArticle: calling ' + model + ' with Q/A length ' + (qaText ? qaText.length : 0));

  var prompt = 'You are a professional newsletter writer and editor.\n'
    + 'Below is a form submission with raw details and responses.\n'
    + 'Analyze the details: if it highlights a person (e.g. name, role, background), write it as an engaging community feature / spotlight; if it describes an event, project milestone, launch, or partnership, write it as an exciting announcement or recap.\n'
    + 'Turn this raw information into a beautifully polished, highly engaging newsletter article suitable for publishing (e.g. on Substack).\n'
    + 'Write in a warm, professional, and appealing tone. Use the answers to build the narrative organically; do not repeat the "Q:" and "A:" labels, template references, or consent checkboxes in the finished article.\n\n'
    + 'Return a JSON object with exactly two fields:\n'
    + '  "title": a short, punchy headline for the finished article (maximum 10 words, no trailing punctuation)\n'
    + '  "body": the full polished article text (no title line, clean paragraph breaks separated by \\n\\n)\n\n'
    + qaText;

  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.7,
      maxOutputTokens: 2100,
      responseMimeType: 'application/json'
    }
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    var responseBody = response.getContentText();

    if (code === 429) {
      Logger.log('GeminiArticle: quota exceeded (429). Retrying once in 36s...');
      Utilities.sleep(36000);
      response = UrlFetchApp.fetch(url, options);
      code = response.getResponseCode();
      responseBody = response.getContentText();
    }

    if (code !== 200) {
      Logger.log('GeminiArticle: API error HTTP ' + code);
      if (code === 429) {
        Logger.log('GeminiArticle: still over quota. Check https://ai.google.dev/gemini-api/docs/rate-limits or try again later.');
      }
      Logger.log('GeminiArticle: response body: ' + responseBody);
      return { title: '', body: '' };
    }

    var data = JSON.parse(responseBody);
    var rawText = data.candidates && data.candidates[0] && data.candidates[0].content
      && data.candidates[0].content.parts && data.candidates[0].content.parts[0]
      ? data.candidates[0].content.parts[0].text
      : '';

    if (!rawText) {
      Logger.log('GeminiArticle: no text in response. candidates: ' + (data.candidates ? data.candidates.length : 0));
      if (data.candidates && data.candidates[0] && data.candidates[0].finishReason) {
        Logger.log('GeminiArticle: finishReason: ' + data.candidates[0].finishReason);
      }
      if (data.error) {
        Logger.log('GeminiArticle: data.error: ' + JSON.stringify(data.error));
      }
      return { title: '', body: '' };
    }

    // Parse the JSON returned by Gemini
    var result;
    try {
      result = JSON.parse(rawText);
    } catch (parseErr) {
      Logger.log('GeminiArticle: JSON parse failed (' + parseErr.message + '). Using raw text as body.');
      return { title: '', body: rawText.trim() };
    }

    var title = (result.title || '').trim().replace(/[.!?]+$/, ''); // strip trailing punctuation
    var body  = (result.body  || '').trim();

    Logger.log('GeminiArticle: success — title: "' + title + '", body: ' + body.length + ' chars');
    return { title: title, body: body };

  } catch (err) {
    Logger.log('GeminiArticle: catch ' + err.message);
    Logger.log('GeminiArticle: stack ' + (err.stack || ''));
    return { title: '', body: '' };
  }
}

function getGeminiApiKey() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
}


// =============================================================================
// SECTION 4: DRAFTS LOG
// Writes the draft title, Doc link, and status to the Drafts Log sheet.
// =============================================================================

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
  var rowStatus  = status     || '✅ Done';
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
  sheet.hideColumns(5);
}

/**
 * Developer utility to test logging to the local spreadsheet.
 */
function testLogDraftToSheet() {
  var row = logDraftToSheet('TEST – Universal Newsletter Draft', 'https://docs.google.com/document/d/example/edit');
  Logger.log('Logged test draft at row ' + row + ' in the "Drafts Log" sheet.');
}


// =============================================================================
// SECTION 5: FORM TRIGGER
// Handles form submissions and orchestrates the full pipeline.
// =============================================================================

/**
 * Adds a "Newsletter Pipeline" menu to the spreadsheet toolbar when opened.
 * Vibe coders can use this instead of the Apps Script function dropdown.
 */
function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('📋 Newsletter Pipeline', [
    { name: '▶ Run Pipeline on Last Response', functionName: 'testWithLastResponse' },
    { name: '🔄 Reprocess Failed Drafts',      functionName: 'reprocessFailed'       },
    { name: '🧪 Submit Test Form Data',         functionName: 'seedTestData'          }
  ]);
}

/**
 * Test: run the pipeline using the last filled row in the active sheet.
 * Use the "Newsletter Pipeline" menu or the Apps Script editor.
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
  var values  = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
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
      ss.toast('⚠️ Gemini is unavailable right now. This submission has been saved as Failed in the Drafts Log. Open the "📋 Newsletter Pipeline" menu and click "🔄 Reprocess Failed Drafts" to retry when Gemini is back.', 'AI Unavailable', 10);
      return;
    }

    var docTitle  = generated.title || Config.DEFAULT_TITLE || 'Article draft';
    var draftBody = generated.body;
    Logger.log('testWithLastResponse: title = "' + docTitle + '"');

    var qaDocUrl = createArticleDoc(docTitle, article.body, ' - Q&A');
    var docUrl   = createArticleDoc(docTitle, draftBody, '', article.visualAssets);
    Logger.log('testWithLastResponse: draft=' + docUrl + ' | Q&A=' + qaDocUrl);

    logDraftToSheet(docTitle, docUrl, '✅ Done', lastRow);
    ss.toast('Draft created and logged! Check the Drafts Log and your Drive folder. 🎉');
  } catch (err) {
    Logger.log('testWithLastResponse ERROR: ' + err.message);
    Logger.log(err.stack || err.toString());
    ss.toast('Error: ' + err.message);
  }
}

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
  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var values     = e.values;
  var sheet      = e.range ? e.range.getSheet() : ss.getActiveSheet();
  var formRowNum = e.range ? e.range.rowStart : '';
  var headers    = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

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
      ss.toast('⚠️ Gemini unavailable. Submission saved as Failed — use the Newsletter Pipeline menu to retry.', 'AI Unavailable', 10);
      return;
    }

    var docTitle  = generated.title || Config.DEFAULT_TITLE || 'Article draft';
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

/**
 * Scans the Drafts Log for rows marked '⚠️ FAILED', retrieves the original
 * form response by row number, and re-runs the full pipeline for each one.
 * On success the log row is updated in place. Call from the Newsletter Pipeline menu.
 */
function reprocessFailed() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(Config.DRAFT_LOG_SHEET_NAME || 'Drafts Log');
  if (!logSheet) { ss.toast('No Drafts Log found. Run setupPipeline first.'); return; }

  var lastRow = logSheet.getLastRow();
  if (lastRow < 2) { ss.toast('The Drafts Log is empty.'); return; }

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

  var formSheet = _getFormResponseSheet(ss);
  if (!formSheet) {
    ss.toast('Could not find the Form Responses sheet. Make sure the form is still linked.');
    return;
  }

  var headers      = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];
  var successCount = 0;
  var failCount    = 0;

  ss.toast('Reprocessing ' + failedEntries.length + ' failed draft(s)… this may take a minute.');

  for (var k = 0; k < failedEntries.length; k++) {
    var entry      = failedEntries[k];
    var formRowNum = entry.formRowNum;
    var logRow     = entry.logRow;

    if (!formRowNum || formRowNum < 2) { failCount++; continue; }

    try {
      var values  = formSheet.getRange(formRowNum, 1, 1, formSheet.getLastColumn()).getValues()[0];
      var article = buildArticleFromResponse(values, headers);
      if (!article || !article.body) { failCount++; continue; }

      var generated = generateArticle(article.body);
      if (!generated.body) { failCount++; continue; }

      var docTitle  = generated.title || Config.DEFAULT_TITLE || 'Article draft';
      var draftBody = generated.body;

      createArticleDoc(docTitle, article.body, ' - Q&A');
      var docUrl = createArticleDoc(docTitle, draftBody, '', article.visualAssets);

      logSheet.getRange(logRow, 2).setValue(docTitle);
      logSheet.getRange(logRow, 3).setValue(docUrl);
      logSheet.getRange(logRow, 4).setValue('✅ Done');

      successCount++;
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


// =============================================================================
// SECTION 6: SETUP
// Run setupPipeline() once to build the form, log sheet, and trigger automatically.
// =============================================================================

/**
 * 1-Click Form & Sheet Initializer.
 * Run this function from the Apps Script editor to automatically build your form,
 * link it to this sheet, set up the Drafts Log, and configure the trigger.
 */
function setupPipeline() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Rename the spreadsheet so it's no longer 'Untitled'
  ss.rename(Config.SPREADSHEET_NAME || 'Newsletter Pipeline');

  // 1. Initialize the Drafts Log on the main (first) sheet FIRST — before linking
  //    the form. When form.setDestination() runs below, Google appends a new
  //    "Form Responses 1" tab, so we must claim Sheet1 before that happens.
  var logSheetName = Config.DRAFT_LOG_SHEET_NAME || 'Drafts Log';
  var logSheet = ss.getSheetByName(logSheetName);
  if (!logSheet) {
    logSheet = ss.getSheets()[0];
    logSheet.setName(logSheetName);
    _initLogHeaders(logSheet);
  }

  // 2. Create the Google Form
  var form = FormApp.create('Newsletter Submission Form');
  form.setDescription('Want to share your story with the community? We\'d love to feature your news, achievements, or insights in our monthly newsletter!');

  // Add the submission questions
  form.addTextItem().setTitle('Your Name').setRequired(true);

  form.addTextItem()
      .setTitle('Your Role/Title')
      .setHelpText('Example: Lead Developer, Community Member, Partner Organization')
      .setRequired(true);

  form.addTextItem().setTitle('Your Telegram').setRequired(true);
  form.addTextItem().setTitle('Your Twitter').setRequired(true);

  form.addParagraphTextItem()
      .setTitle('Summary (2-4 sentences)')
      .setHelpText('Give us the core message. What\'s the key takeaway?')
      .setRequired(true);

  form.addParagraphTextItem()
      .setTitle('Full Description')
      .setHelpText('Share the complete story, details, background, or context. What impact does it have? What do you want readers to do?')
      .setRequired(true);

  form.addParagraphTextItem()
      .setTitle('Links')
      .setHelpText('Relevant URLs: websites, social posts, articles, registration pages, etc.');

  form.addParagraphTextItem()
      .setTitle('Visual Assets')
      .setHelpText('Share photos, logos, graphics via link (Google Drive, Dropbox, etc.)');

  form.addParagraphTextItem()
      .setTitle('Contact Information (if applicable)')
      .setHelpText('(Email, social handles, website)');

  var consent = form.addCheckboxItem();
  consent.setTitle('Consent & Attribution')
         .setChoices([
           consent.createChoice('I have permission to submit this content and any included photos/quotes'),
           consent.createChoice('I consent to editing this content for clarity, length, and style')
         ])
         .setRequired(true);

  // 3. Link the form — this appends a new "Form Responses 1" tab after the log
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // Move the Form into the same Drive folder as this spreadsheet
  var ssParents = DriveApp.getFileById(ss.getId()).getParents();
  if (ssParents.hasNext()) {
    DriveApp.getFileById(form.getId()).moveTo(ssParents.next());
  }

  // 3. Programmatically install the On Form Submit trigger
  var existingTriggers = ScriptApp.getProjectTriggers();
  var triggerExists = false;
  for (var i = 0; i < existingTriggers.length; i++) {
    if (existingTriggers[i].getHandlerFunction() === 'onFormSubmit') {
      triggerExists = true;
      break;
    }
  }

  if (!triggerExists) {
    ScriptApp.newTrigger('onFormSubmit')
             .forSpreadsheet(ss)
             .onFormSubmit()
             .create();
    Logger.log('Installed form submit trigger successfully.');
  } else {
    Logger.log('Form submit trigger already exists.');
  }

  // Output setup completion details
  Logger.log('====================================================');
  Logger.log('🎉 PIPELINE SETUP SUCCESSFUL!');
  Logger.log('📝 Edit your new Google Form here: ' + form.getEditUrl());
  Logger.log('🔗 Share this link with your submitters: ' + form.getPublishedUrl());
  Logger.log('====================================================');

  ss.toast('Pipeline initialized! Check the execution log for Form URLs.', 'Success 🎉');
}


// =============================================================================
// SECTION 7: TEST DATA SEEDER
// Uses Gemini to generate 3 diverse, realistic test submissions on the spot,
// then submits them through the linked form so the full pipeline runs end-to-end.
// Call from the "Newsletter Pipeline" menu → "🧪 Submit Test Form Data".
// =============================================================================

/**
 * Asks Gemini to generate 3 diverse test form submissions, then submits each
 * one through the linked Google Form. The onFormSubmit trigger fires for each,
 * running the full AI pipeline automatically.
 *
 * Set Config.ORG_NAME to tailor the content to your organization.
 * Leave it blank for generic output that works with any org.
 */
function seedTestData() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var formUrl = ss.getFormUrl();

  if (!formUrl) {
    ss.toast('No form is linked to this spreadsheet. Run setupPipeline first, then try again.');
    return;
  }

  ss.toast('Asking Gemini to generate test submissions… this takes a few seconds.');

  var submissions = _generateTestSubmissions();
  if (!submissions || submissions.length === 0) {
    ss.toast('Could not generate test data — check that your GEMINI_API_KEY is set and try again.');
    return;
  }

  var form  = FormApp.openByUrl(formUrl);
  var count = 0;

  for (var i = 0; i < submissions.length; i++) {
    try {
      submissions[i]['Consent & Attribution'] = [
        'I have permission to submit this content and any included photos/quotes',
        'I consent to editing this content for clarity, length, and style'
      ];
      _submitFormResponse(form, submissions[i]);
      count++;
      Utilities.sleep(1500);
    } catch (err) {
      Logger.log('seedTestData: error on submission ' + (i + 1) + ': ' + err.message);
    }
  }

  ss.toast(count + ' of ' + submissions.length + ' AI-generated test responses submitted! Check your Drafts Log and Drive folder in a moment. 🎉', 'Test Data Seeded');
}

/**
 * Calls Gemini to generate an array of 3 diverse, realistic newsletter form
 * submissions. Uses Config.ORG_NAME when set for org-aware content.
 *
 * @return {Array.<Object>|null} array of answer objects, or null on failure
 */
function _generateTestSubmissions() {
  var apiKey = getGeminiApiKey();
  if (!apiKey) { Logger.log('_generateTestSubmissions: GEMINI_API_KEY not set.'); return null; }

  var orgContext = Config.ORG_NAME
    ? 'The organization is called "' + Config.ORG_NAME + '".'
    : 'Use a fictional but realistic organization name — do NOT use placeholder text like "Your Organization".';

  var model = Config.GEMINI_MODEL || 'gemini-2.5-flash-lite';
  var url   = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent';

  var prompt =
    'You are helping test a newsletter submission pipeline. Generate 3 realistic, engaging, and diverse newsletter form submissions.\n'
    + orgContext + '\n'
    + 'Cover these three submission types (one each):\n'
    + '  1. Member Spotlight — highlights a specific person\'s journey, contributions, or achievement\n'
    + '  2. Event Recap — recaps a community event: hackathon, workshop, meetup, or conference\n'
    + '  3. Announcement — a partnership, launch, grant, or initiative\n\n'
    + 'Return a JSON array of exactly 3 objects. Each object must have these exact keys:\n'
    + '  "Your Name" — submitter\'s full name\n'
    + '  "Your Role/Title" — their role (can reference the org name)\n'
    + '  "Your Telegram" — Telegram handle starting with @\n'
    + '  "Your Twitter" — Twitter/X handle starting with @\n'
    + '  "Summary (2-4 sentences)" — 2–4 compelling sentences summarizing the story\n'
    + '  "Full Description" — 3–5 detailed paragraphs telling the full story\n'
    + '  "Links" — 1–3 relevant URLs separated by | characters\n'
    + '  "Visual Assets" — a placeholder Google Drive folder URL\n'
    + '  "Contact Information (if applicable)" — email and/or social handles\n\n'
    + 'Make the content vivid, specific, and publishable. Vary tone and subject matter across the three submissions.';

  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.9, maxOutputTokens: 3000, responseMimeType: 'application/json' }
  };
  var options = {
    method: 'post', contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify(payload), muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    var body = response.getContentText();

    if (code === 429 || code === 503) {
      var waitMs = code === 429 ? 36000 : 10000;
      Logger.log('_generateTestSubmissions: error ' + code + '. Retrying in ' + (waitMs / 1000) + 's...');
      Utilities.sleep(waitMs);
      response = UrlFetchApp.fetch(url, options);
      code = response.getResponseCode();
      body = response.getContentText();
    }

    if (code !== 200) { Logger.log('_generateTestSubmissions: API error ' + code + ': ' + body); return null; }

    var data    = JSON.parse(body);
    var rawText = data.candidates && data.candidates[0] && data.candidates[0].content
      && data.candidates[0].content.parts && data.candidates[0].content.parts[0]
      ? data.candidates[0].content.parts[0].text : '';

    if (!rawText) { Logger.log('_generateTestSubmissions: empty response'); return null; }

    var submissions = JSON.parse(rawText);
    Logger.log('_generateTestSubmissions: generated ' + submissions.length + ' submissions');
    return Array.isArray(submissions) ? submissions : null;

  } catch (err) {
    Logger.log('_generateTestSubmissions: error — ' + err.message);
    return null;
  }
}

/**
 * Builds and submits one Google Form response from a plain answers object.
 * Keys must match form item titles exactly.
 * Checkbox answers should be passed as an array of choice strings.
 *
 * @param {GoogleAppsScript.Forms.Form} form
 * @param {Object} answers - { 'Question title': 'Answer text' | ['Choice 1', ...] }
 */
function _submitFormResponse(form, answers) {
  var items    = form.getItems();
  var response = form.createResponse();

  for (var i = 0; i < items.length; i++) {
    var item   = items[i];
    var title  = item.getTitle();
    var answer = answers[title];
    if (answer === undefined || answer === null || answer === '') continue;

    var itemType     = item.getType();
    var itemResponse = null;

    if (itemType === FormApp.ItemType.TEXT) {
      itemResponse = item.asTextItem().createResponse(String(answer));
    } else if (itemType === FormApp.ItemType.PARAGRAPH_TEXT) {
      itemResponse = item.asParagraphTextItem().createResponse(String(answer));
    } else if (itemType === FormApp.ItemType.CHECKBOX) {
      var choices = Array.isArray(answer) ? answer : [String(answer)];
      itemResponse = item.asCheckboxItem().createResponse(choices);
    }

    if (itemResponse) response.withItemResponse(itemResponse);
  }

  response.submit();
}

/**
 * Asks Gemini to generate 3 diverse test form submissions, then submits each
 * one through the linked Google Form. The onFormSubmit trigger fires for each,
 * running the full AI pipeline automatically.
 *
 * Set Config.ORG_NAME to tailor the content to your organization.
 * Leave it blank for generic output that works with any org.
 */
function seedTestData() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var formUrl = ss.getFormUrl();

  if (!formUrl) {
    ss.toast('No form is linked to this spreadsheet. Run setupPipeline first, then try again.', 'Setup Required', 8);
    return;
  }

  ss.toast('Asking Gemini to generate test submissions… this may take up to 30 seconds if the API is busy.', '🧪 Generating…', 30);

  var result = _generateTestSubmissions();

  if (!result.data) {
    var msg;
    switch (result.error) {
      case 'NO_KEY':
        msg = 'GEMINI_API_KEY is not set. Go to Apps Script → Project Settings → Script Properties and add your key, then try again.';
        break;
      case 'QUOTA':
        msg = 'Gemini API quota exceeded. Wait 1–2 minutes and try again, or check your usage at aistudio.google.com.';
        break;
      case 'UNAVAILABLE':
        msg = 'Gemini is experiencing high demand right now. All 3 retry attempts failed. Try again in a few minutes.';
        break;
      case 'PARSE':
        msg = 'Gemini returned unexpected data. Try again — it usually works on the next attempt.';
        break;
      default:
        msg = 'Could not generate test data (' + (result.error || 'unknown error') + '). Check the execution log for details.';
    }
    ss.toast(msg, '⚠️ Test Data Failed', 12);
    return;
  }

  var submissions = result.data;
  var form  = FormApp.openByUrl(formUrl);
  var count = 0;

  for (var i = 0; i < submissions.length; i++) {
    try {
      // Attach consent checkboxes — same for every submission
      submissions[i]['Consent & Attribution'] = [
        'I have permission to submit this content and any included photos/quotes',
        'I consent to editing this content for clarity, length, and style'
      ];
      _submitFormResponse(form, submissions[i]);
      count++;
      Utilities.sleep(1500);
    } catch (err) {
      Logger.log('seedTestData: error on submission ' + (i + 1) + ': ' + err.message);
    }
  }

  if (count === 0) {
    ss.toast('Submissions were generated but none could be posted to the form. Check the execution log.', '⚠️ Submit Failed', 10);
  } else if (count < submissions.length) {
    ss.toast(count + ' of ' + submissions.length + ' test responses submitted. Some failed — check the execution log.', 'Partial Success', 8);
  } else {
    ss.toast(count + ' AI-generated test responses submitted! Check your Drafts Log and Drive folder in a moment. 🎉', 'Test Data Seeded', 8);
  }
}

/**
 * Calls Gemini to generate an array of 3 diverse, realistic newsletter form
 * submissions. Retries up to 3 times total with exponential backoff.
 *
 * @return {{ data: Array.<Object>|null, error: string|null }}
 *   data  — array of submission objects on success, null on failure
 *   error — 'NO_KEY' | 'QUOTA' | 'UNAVAILABLE' | 'PARSE' | 'NETWORK' | null
 */
function _generateTestSubmissions() {
  var apiKey = getGeminiApiKey();
  if (!apiKey) {
    Logger.log('_generateTestSubmissions: GEMINI_API_KEY not set.');
    return { data: null, error: 'NO_KEY' };
  }

  var orgContext = Config.ORG_NAME
    ? 'The organization is called "' + Config.ORG_NAME + '".'
    : 'Use a fictional but realistic organization name — do NOT use placeholder text like "Your Organization".';

  var model = Config.GEMINI_MODEL || 'gemini-2.5-flash-lite';
  var url   = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent';

  var prompt =
    'You are helping test a newsletter submission pipeline. Generate 3 realistic, engaging, and diverse newsletter form submissions.\n'
    + orgContext + '\n'
    + 'Cover these three submission types (one each):\n'
    + '  1. Member Spotlight — highlights a specific person\'s journey, contributions, or achievement\n'
    + '  2. Event Recap — recaps a community event: hackathon, workshop, meetup, or conference\n'
    + '  3. Announcement — a partnership, launch, grant, or initiative\n\n'
    + 'Return a JSON array of exactly 3 objects. Each object must have these exact keys:\n'
    + '  "Your Name" — submitter\'s full name\n'
    + '  "Your Role/Title" — their role (can reference the org name)\n'
    + '  "Your Telegram" — Telegram handle starting with @\n'
    + '  "Your Twitter" — Twitter/X handle starting with @\n'
    + '  "Summary (2-4 sentences)" — 2–4 compelling sentences summarizing the story\n'
    + '  "Full Description" — 3–5 detailed paragraphs telling the full story\n'
    + '  "Links" — 1–3 relevant URLs separated by | characters\n'
    + '  "Visual Assets" — a placeholder Google Drive folder URL\n'
    + '  "Contact Information (if applicable)" — email and/or social handles\n\n'
    + 'Make the content vivid, specific, and publishable. Vary tone and subject matter across the three submissions.';

  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.9, maxOutputTokens: 3000, responseMimeType: 'application/json' }
  };
  var options = {
    method: 'post', contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify(payload), muteHttpExceptions: true
  };

  // Retry schedule: [wait before attempt 2, wait before attempt 3]
  // 503 (high demand): 10s → 20s   429 (quota): 36s → 72s
  var MAX_ATTEMPTS = 3;
  var lastCode = 0;
  var lastBody = '';

  for (var attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    try {
      var response = UrlFetchApp.fetch(url, options);
      lastCode = response.getResponseCode();
      lastBody = response.getContentText();

      if (lastCode === 200) {
        var data    = JSON.parse(lastBody);
        var rawText = data.candidates && data.candidates[0] && data.candidates[0].content
          && data.candidates[0].content.parts && data.candidates[0].content.parts[0]
          ? data.candidates[0].content.parts[0].text : '';

        if (!rawText) {
          Logger.log('_generateTestSubmissions: empty text in response (attempt ' + attempt + ')');
          return { data: null, error: 'PARSE' };
        }

        var submissions;
        try { submissions = JSON.parse(rawText); }
        catch (pe) {
          Logger.log('_generateTestSubmissions: JSON parse error — ' + pe.message);
          return { data: null, error: 'PARSE' };
        }

        if (!Array.isArray(submissions) || submissions.length === 0) {
          Logger.log('_generateTestSubmissions: response was not a non-empty array');
          return { data: null, error: 'PARSE' };
        }

        Logger.log('_generateTestSubmissions: success on attempt ' + attempt + ' — ' + submissions.length + ' submissions');
        return { data: submissions, error: null };
      }

      if ((lastCode === 429 || lastCode === 503) && attempt < MAX_ATTEMPTS) {
        var waitMs = lastCode === 429
          ? (attempt === 1 ? 36000 : 72000)
          : (attempt === 1 ? 10000 : 20000);
        Logger.log('_generateTestSubmissions: HTTP ' + lastCode + ' on attempt ' + attempt + '. Waiting ' + (waitMs / 1000) + 's before retry…');
        Utilities.sleep(waitMs);
        continue;
      }

      Logger.log('_generateTestSubmissions: HTTP ' + lastCode + ' on attempt ' + attempt + ' (final).');
      break;

    } catch (err) {
      Logger.log('_generateTestSubmissions: network error on attempt ' + attempt + ' — ' + err.message);
      lastCode = 0;
      if (attempt < MAX_ATTEMPTS) { Utilities.sleep(5000); continue; }
      return { data: null, error: 'NETWORK' };
    }
  }

  var errorType = lastCode === 429 ? 'QUOTA' : 'UNAVAILABLE';
  return { data: null, error: errorType };
}

/**
 * Builds and submits one Google Form response from a plain answers object.
 * Keys must match form item titles exactly.
 * Checkbox answers should be passed as an array of choice strings.
 *
 * @param {GoogleAppsScript.Forms.Form} form
 * @param {Object} answers - { 'Question title': 'Answer text' | ['Choice 1', ...] }
 */
function _submitFormResponse(form, answers) {
  var items    = form.getItems();
  var response = form.createResponse();

  for (var i = 0; i < items.length; i++) {
    var item   = items[i];
    var title  = item.getTitle();
    var answer = answers[title];
    if (answer === undefined || answer === null || answer === '') continue;

    var itemType     = item.getType();
    var itemResponse = null;

    if (itemType === FormApp.ItemType.TEXT) {
      itemResponse = item.asTextItem().createResponse(String(answer));
    } else if (itemType === FormApp.ItemType.PARAGRAPH_TEXT) {
      itemResponse = item.asParagraphTextItem().createResponse(String(answer));
    } else if (itemType === FormApp.ItemType.CHECKBOX) {
      var choices = Array.isArray(answer) ? answer : [String(answer)];
      itemResponse = item.asCheckboxItem().createResponse(choices);
    }

    if (itemResponse) response.withItemResponse(itemResponse);
  }

  response.submit();
}

