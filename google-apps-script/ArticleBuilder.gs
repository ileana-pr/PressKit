/**
 * Builds Q/A article text from one form response.
 * Title generation is handled by Gemini (generateTitleFromQa in GeminiArticle.gs).
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
