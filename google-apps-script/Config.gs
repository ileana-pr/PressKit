/**
 * Master Newsletter Pipeline – General Configuration
 * Universal template for any organization.
 */
var Config = {
  // Name of the sheet/tab in this spreadsheet where generated drafts will be logged
  DRAFT_LOG_SHEET_NAME: 'Drafts Log',

  // Default fallback if a title cannot be found
  DEFAULT_TITLE: 'Untitled Draft',

  // New draft Doc: optional prefix in file name. Date (YYYY-MM-DD) is always prepended.
  DOC_TITLE_PREFIX: '',

  // Folder ID where new draft Google Docs will be created. Leave blank to save in Google Drive root.
  DRAFT_FOLDER_ID: '1anKyPLe5k11wxnRa8tSrB-q-wmEsFZ9s',

  // Gemini model for content generation.
  // gemini-2.5-flash-lite is fast, highly capable, and has generous free-tier quotas.
  GEMINI_MODEL: 'gemini-2.5-flash-lite'
};
