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
// Submits realistic sample responses directly through the linked Google Form
// so you can test the full pipeline without filling in the form manually.
// Call from the "Newsletter Pipeline" menu → "Submit Test Form Data".
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
      Utilities.sleep(1500); // small pause between submissions
    } catch (err) {
      Logger.log('seedTestData: error on submission ' + (i + 1) + ': ' + err.message);
    }
  }

  ss.toast(count + ' of ' + TEST_SUBMISSIONS.length + ' test responses submitted! The pipeline trigger fires for each one — check your Drafts Log and Drive folder in a moment. 🎉', 'Test Data Seeded');
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
      // answer must be an array of choice strings
      var choices = Array.isArray(answer) ? answer : [String(answer)];
      itemResponse = item.asCheckboxItem().createResponse(choices);
    }

    if (itemResponse) response.withItemResponse(itemResponse);
  }

  response.submit();
}
