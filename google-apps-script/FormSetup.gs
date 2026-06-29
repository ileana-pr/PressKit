/**
 * 1-Click Form & Sheet Initializer
 * Run the setupPipeline() function from this file to automatically build your form,
 * link it to this sheet, create the Drafts Log, and configure the triggers!
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

  // 4. Programmatically install the On Form Submit trigger
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
// TEST DATA SEEDER
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

  var model = Config.GEMINI_MODEL || 'gemini-2.0-flash';
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
  // 503 (high demand): 10s then 20s — temporary spikes usually clear quickly
  // 429 (quota):       36s then 72s — quota windows reset on a per-minute basis
  var MAX_ATTEMPTS = 3;
  var lastCode = 0;
  var lastBody = '';

  for (var attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    try {
      var response = UrlFetchApp.fetch(url, options);
      lastCode = response.getResponseCode();
      lastBody = response.getContentText();

      if (lastCode === 200) {
        // ── Success path ──
        var data    = JSON.parse(lastBody);
        var rawText = data.candidates && data.candidates[0] && data.candidates[0].content
          && data.candidates[0].content.parts && data.candidates[0].content.parts[0]
          ? data.candidates[0].content.parts[0].text : '';

        if (!rawText) {
          Logger.log('_generateTestSubmissions: empty text in response (attempt ' + attempt + ')');
          return { data: null, error: 'PARSE' };
        }

        var submissions;
        try {
          submissions = JSON.parse(rawText);
        } catch (parseErr) {
          Logger.log('_generateTestSubmissions: JSON parse error — ' + parseErr.message);
          return { data: null, error: 'PARSE' };
        }

        if (!Array.isArray(submissions) || submissions.length === 0) {
          Logger.log('_generateTestSubmissions: response was not a non-empty array');
          return { data: null, error: 'PARSE' };
        }

        Logger.log('_generateTestSubmissions: success on attempt ' + attempt + ' — ' + submissions.length + ' submissions');
        return { data: submissions, error: null };
      }

      // ── Retryable errors ──
      if ((lastCode === 429 || lastCode === 503) && attempt < MAX_ATTEMPTS) {
        var waitMs = lastCode === 429
          ? (attempt === 1 ? 36000 : 72000)   // 36s, 72s
          : (attempt === 1 ? 10000 : 20000);   // 10s, 20s
        Logger.log('_generateTestSubmissions: HTTP ' + lastCode + ' on attempt ' + attempt + '. Waiting ' + (waitMs / 1000) + 's before retry…');
        Utilities.sleep(waitMs);
        continue; // next attempt
      }

      // ── Non-retryable or exhausted retries ──
      Logger.log('_generateTestSubmissions: HTTP ' + lastCode + ' on attempt ' + attempt + ' (final). Body: ' + lastBody);
      break;

    } catch (err) {
      Logger.log('_generateTestSubmissions: network error on attempt ' + attempt + ' — ' + err.message);
      lastCode = 0;
      if (attempt < MAX_ATTEMPTS) {
        Utilities.sleep(5000);
        continue;
      }
      return { data: null, error: 'NETWORK' };
    }
  }

  // Classify the final error code for a targeted toast message
  var errorType = lastCode === 429 ? 'QUOTA' : lastCode === 503 ? 'UNAVAILABLE' : 'UNAVAILABLE';
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


