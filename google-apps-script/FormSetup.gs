/**
 * 1-Click Form & Sheet Initializer
 * Run the setupPipeline() function from this file to automatically build your form,
 * link it to this sheet, create the Drafts Log, and configure the triggers!
 */
function setupPipeline() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Initialize the Drafts Log on the main (first) sheet FIRST — before linking
  //    the form. When form.setDestination() runs below, Google appends a new
  //    "Form Responses 1" tab, so we must claim Sheet1 before that happens.
  var logSheetName = Config.DRAFT_LOG_SHEET_NAME || 'Drafts Log';
  var logSheet = ss.getSheetByName(logSheetName);
  if (!logSheet) {
    logSheet = ss.getSheets()[0];
    logSheet.setName(logSheetName);
    logSheet.appendRow(['Date', 'Draft Title', 'Google Doc Link']);
    logSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    logSheet.setColumnWidth(1, 120);
    logSheet.setColumnWidth(2, 300);
    logSheet.setColumnWidth(3, 400);
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
