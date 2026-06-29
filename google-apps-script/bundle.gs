// =============================================================================
// HER DAO Newsletter Pipeline — Single-File Bundle
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
// Submits realistic sample responses directly through the linked Google Form
// so you can test the full pipeline without filling in the form manually.
// Call from the "Newsletter Pipeline" menu → "🧪 Submit Test Form Data".
// =============================================================================

var TEST_SUBMISSIONS = [
  {
    'Your Name': 'Valentina Cruz',
    'Your Role/Title': 'Smart Contract Developer & Web3 Educator',
    'Your Telegram': '@valcruzcodes',
    'Your Twitter': '@valcruzcodes',
    'Summary (2-4 sentences)': 'Valentina Cruz joined HER DAO eight months ago as a curious newcomer to Web3. Since then, she has shipped two community tools, mentored five junior devs, and become one of the most active voices in the governance forum. Her journey from "what is a wallet?" to writing Solidity in production is exactly the kind of story we love to share.',
    'Full Description': 'Valentina came from a background in traditional fintech, where she spent four years building payment APIs. A Twitter thread about DAOs sent her down a rabbit hole in early 2024, and she signed up for HER DAO\'s onboarding cohort within weeks. She credits the DAO\'s mentorship structure as the thing that kept her going when Solidity felt impossible. Her first contribution was a small script to automate the weekly governance recap — it saved the ops team roughly three hours every week. She has since built a public dashboard that tracks HER DAO\'s treasury in real time, which is now used by the core team in every budget call. Outside of code, Valentina runs a free weekly "Web3 Wednesdays" session for Spanish-speaking developers, pulling in 30–50 attendees each week. She believes that the biggest gap in the ecosystem is not technical — it\'s linguistic and cultural, and she intends to keep building bridges.',
    'Links': 'https://valentina.dev | https://twitter.com/valcruzcodes | https://github.com/valcruzcodes',
    'Visual Assets': 'https://drive.google.com/drive/folders/example-valentina-photos',
    'Contact Information (if applicable)': 'valentina@valentina.dev | @valcruzcodes on Telegram',
    'Consent & Attribution': ['I have permission to submit this content and any included photos/quotes', 'I consent to editing this content for clarity, length, and style']
  },
  {
    'Your Name': 'Priya Mehta',
    'Your Role/Title': 'Community Lead, HER DAO',
    'Your Telegram': '@priyaherdao',
    'Your Twitter': '@priya_herdao',
    'Summary (2-4 sentences)': 'HER DAO hosted its first in-person hackathon in Lisbon on June 14–15, bringing together 60 builders from 18 countries. Teams had 24 hours to build on the theme of "financial sovereignty for underbanked communities." The winning project — a mobile-first micro-savings protocol — earned a $5,000 grant and is already in talks with two DeFi protocols for integration.',
    'Full Description': 'The Lisbon Hackathon was over a year in the making, born from a community poll asking members what they wanted most: more online events or one big in-person moment. The vote wasn\'t even close. Sixty participants flew in from across Europe, Latin America, and West Africa, many meeting their Discord friends IRL for the very first time. The venue — a converted tile factory in the LX Factory complex — was packed with laptops, whiteboards, and the occasional burst of collective cheering when a demo finally worked. Twelve teams submitted projects covering on-chain identity for undocumented workers and DAO tooling built specifically for mobile-first markets. The winning team, FloraSave, built a micro-savings app that rounds up crypto transactions and pools the change into a community yield vault — no bank account required. The post-event survey showed a 96% satisfaction rate, and the community is already voting on whether to make Lisbon an annual tradition.',
    'Links': 'https://herdao.xyz/hackathon-lisbon | https://devpost.com/herdao-lisbon',
    'Visual Assets': 'https://drive.google.com/drive/folders/example-lisbon-hackathon-photos',
    'Contact Information (if applicable)': 'priya@herdao.xyz',
    'Consent & Attribution': ['I have permission to submit this content and any included photos/quotes', 'I consent to editing this content for clarity, length, and style']
  },
  {
    'Your Name': 'Amara Osei',
    'Your Role/Title': 'Partnerships Lead, HER DAO',
    'Your Telegram': '@amaraherdao',
    'Your Twitter': '@amara_osei',
    'Summary (2-4 sentences)': 'HER DAO has partnered with Stellar Development Foundation to fund 10 grants of $3,000 each for women-led Web3 projects in Sub-Saharan Africa. Applications open July 1st and close July 31st. Projects must be in early or growth stage and led by at least one woman founder with ties to the region.',
    'Full Description': 'This partnership has been in conversations since ETHDenver, when HER DAO and SDF representatives spent an afternoon swapping notes on what gaps they were each seeing in the ecosystem. The answer kept coming back to the same thing: brilliant founders in Sub-Saharan Africa who lacked the runway to turn their prototypes into products. Ten teams will each receive $3,000 in funding, plus six months of mentorship through the HER DAO network, access to SDF\'s technical resources, and a guaranteed spot to present at HER DAO\'s virtual demo day in October. Applications are intentionally kept short — a two-page summary and a three-minute video pitch — because long applications systematically disadvantage founders who are juggling everything. The selection committee includes judges from HER DAO, SDF, and independent experts in African fintech. Winners will be announced August 15th.',
    'Links': 'https://herdao.xyz/sdf-grant | https://stellar.org/foundation',
    'Visual Assets': 'https://drive.google.com/drive/folders/example-sdf-partnership-assets',
    'Contact Information (if applicable)': 'grants@herdao.xyz | @amaraherdao on Telegram',
    'Consent & Attribution': ['I have permission to submit this content and any included photos/quotes', 'I consent to editing this content for clarity, length, and style']
  }
];

/**
 * Submits all TEST_SUBMISSIONS programmatically through the linked Google Form.
 * Each submission fires the onFormSubmit trigger and runs the full AI pipeline.
 * Call from the "Newsletter Pipeline" menu → "🧪 Submit Test Form Data".
 */
function seedTestData() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var formUrl = ss.getFormUrl();

  if (!formUrl) {
    ss.toast('No form is linked to this spreadsheet. Run setupPipeline first, then try again.');
    return;
  }

  var form  = FormApp.openByUrl(formUrl);
  var count = 0;

  ss.toast('Submitting ' + TEST_SUBMISSIONS.length + ' test responses — the pipeline will run automatically for each one…');

  for (var i = 0; i < TEST_SUBMISSIONS.length; i++) {
    try {
      _submitFormResponse(form, TEST_SUBMISSIONS[i]);
      count++;
      Utilities.sleep(1500);
    } catch (err) {
      Logger.log('seedTestData: error on submission ' + (i + 1) + ': ' + err.message);
    }
  }

  ss.toast(count + ' of ' + TEST_SUBMISSIONS.length + ' test responses submitted! Check your Drafts Log and Drive folder in a moment. 🎉', 'Test Data Seeded');
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
