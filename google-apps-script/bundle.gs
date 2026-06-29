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
  // Name of the sheet/tab in this spreadsheet where generated drafts will be logged
  DRAFT_LOG_SHEET_NAME: 'Drafts Log',

  // Default fallback title if AI generation fails
  DEFAULT_TITLE: 'Untitled Draft',

  // New draft Doc: optional prefix in file name. Date (YYYY-MM-DD) is always prepended.
  DOC_TITLE_PREFIX: '',

  // Folder ID where new draft Google Docs will be created. Leave blank ('') to save in Drive root.
  // Find the ID in your folder's URL: drive.google.com/drive/folders/YOUR_ID_HERE
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

  if (Config.DRAFT_FOLDER_ID) {
    var file = DriveApp.getFileById(doc.getId());
    file.moveTo(DriveApp.getFolderById(Config.DRAFT_FOLDER_ID));
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
// Writes the draft title and Doc link to the Drafts Log sheet.
// =============================================================================

/**
 * Logs a generated draft directly to the Drafts Log sheet in this spreadsheet.
 * If the sheet doesn't exist, it automatically creates it with headers.
 *
 * @param {string} title - title of the draft
 * @param {string} docUrl - link to the generated Google Doc
 * @return {number} row index of the logged draft
 */
function logDraftToSheet(title, docUrl) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = Config.DRAFT_LOG_SHEET_NAME || 'Drafts Log';
  var sheet = ss.getSheetByName(sheetName);

  // If the log sheet doesn't exist, create it with headers
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['Date', 'Draft Title', 'Google Doc Link']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 300);
    sheet.setColumnWidth(3, 400);
  }

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([today, title, docUrl]);
  return sheet.getLastRow();
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

  // 1. Create the Google Form
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

  // Link the new Form to this spreadsheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // 2. Initialize the Drafts Log on the main (first) sheet.
  // When form.setDestination() runs above, Google creates a new "Form Responses 1"
  // tab automatically — so the original first sheet stays empty and is the perfect
  // home for the log. We rename it and write headers there.
  var logSheetName = Config.DRAFT_LOG_SHEET_NAME || 'Drafts Log';
  var logSheet = ss.getSheetByName(logSheetName);
  if (!logSheet) {
    // Use the first sheet (empty) rather than inserting a new tab
    logSheet = ss.getSheets()[0];
    logSheet.setName(logSheetName);
    logSheet.appendRow(['Date', 'Draft Title', 'Google Doc Link']);
    logSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    logSheet.setColumnWidth(1, 120);
    logSheet.setColumnWidth(2, 300);
    logSheet.setColumnWidth(3, 400);
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
