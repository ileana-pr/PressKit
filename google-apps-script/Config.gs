/**
 * PressKit — AI Newsletter Pipeline · General Configuration
 * Universal template for any organization.
 */
var Config = {
  // (Optional) Your organization's name.
  // When set, Gemini will tailor generated test submissions and article tone to your org.
  // Leave blank ('') for generic output that works with any organization.
  ORG_NAME: '',

  // Name given to the spreadsheet when setupPipeline() runs
  SPREADSHEET_NAME: 'PressKit',

  // Name of the sheet/tab in this spreadsheet where generated drafts will be logged
  DRAFT_LOG_SHEET_NAME: 'Drafts Log',

  // Default fallback if a title cannot be found
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
