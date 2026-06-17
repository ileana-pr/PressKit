# 🌟 Universal AI Newsletter Pipeline (Consolidated Template)

An elegant, zero-overhead Google Apps Script automation that turns raw Google Form submissions into fully written, polished newsletter drafts using Gemini AI, and logs them directly inside the same response spreadsheet.

This repository is designed specifically for **vibecoders** and non-technical people who want a premium, friction-free "plug-and-play" pipeline. It features a **1-click automatic setup** that builds your form, links your sheets, and installs triggers in under 3 minutes!

---

## ⚡ The Universal Flow

```
                      [ User Submits Form ]
                  (Raw Story, Details, Links)
                               │
                               ▼
                   [ Google Sheet Response ]
                               │
                               ▼
            [ Apps Script Trigger: onFormSubmit ]
             ├── 1. Dynamic Title Extraction & Truncation
             ├── 2. Automatic Q&A Packaging
             └── 3. Generates Google Doc Draft
                               │
                               ▼
          [ Gemini AI (gemini-2.5-flash-lite) ]
          (Adaptive prompt writes perfectly in any style:
           spotlight, event recap, announcement, etc.)
                               │
                               ▼
              [ Google Doc Created in Drive ]
                               │
                               ▼
         [ Logs Draft Title & Link to "Drafts Log" Tab ]
```

---

## 💎 Why This Pipeline Rules

*   **1-Click Automated Setup:** Run a single helper function and Google will build the 10-question form, connect it to your sheet, create logging tabs, and register triggers programmatically. No manual clicking required!
*   **Only One Spreadsheet Needed:** No secondary tracker spreadsheets to configure. Everything is saved and logged in the same response sheet.
*   **Adaptive Gemini Prompt:** The AI reads the raw submission content, automatically analyzes the style (recognizing if it highlights a person, recaps a past event, shares a new launch, or announces a partnership), and writes the newsletter draft in the perfect style.
*   **Smart Title Truncation:** Automatically identifies the best summary or title column, and cleanly truncates long descriptions to keep Google Doc names and logging sheets looking clean.

---

## 🚀 3-Step Setup Guide

### 1. Copy the Code to Apps Script
1. Create a brand new **Google Sheet**.
2. Go to the top menu and select **Extensions ➔ Apps Script**.
3. Create the following files in the Apps Script editor (using the **+** button next to Files), and copy-paste the code from this repository:
    *   [Config.gs](google-apps-script/Config.gs)
    *   [ArticleBuilder.gs](google-apps-script/ArticleBuilder.gs)
    *   [GeminiArticle.gs](google-apps-script/GeminiArticle.gs)
    *   [FormTrigger.gs](google-apps-script/FormTrigger.gs)
    *   [appendArticleTask.gs](google-apps-script/appendArticleTask.gs)
    *   [FormSetup.gs](google-apps-script/FormSetup.gs)

### 2. Run the 1-Click Initializer
1. In the Apps Script editor, open the file `FormSetup.gs`.
2. In the top toolbar, make sure **`setupPipeline`** is selected in the function dropdown, then click **Run**.
3. **Authorize Permissions:** Follow the prompt to authorize permissions (click *Advanced* ➔ *Go to Untitled Project (unsafe)*).
4. **Copy your Form Links:** Once the script finishes, click the **Execution Log** at the bottom. It will print the links to edit and share your new Google Form!

### 3. Customize Config & Add Gemini API Key
Open your script's [Config.gs](google-apps-script/Config.gs):
1.  **`DRAFT_FOLDER_ID`:** Replace with the ID of your Google Drive folder where new article draft documents should be saved (found in the folder's URL: `drive.google.com/drive/folders/YOUR_ID_HERE`). *Note: Leave blank to save directly in the root of your Google Drive.*
2.  **Add Gemini API Key:**
    *   Get a free API Key from [Google AI Studio](https://aistudio.google.com/).
    *   In Apps Script, click the **Project Settings (gear)** icon ➔ Scroll down to **Script Properties** ➔ Click **Add script property**.
    *   Set the Property to `GEMINI_API_KEY` and paste your API key as the Value. Save properties.

---

### 🎨 The Test Run (Watch the Magic!)
1. Open the shareable Google Form link you copied in Step 2.
2. Submit a test response to the form.
3. Go back to your Apps Script editor, open `FormTrigger.gs`, select `testWithLastResponse` in the top dropdown, and click **Run**.
4. A new **"Drafts Log"** tab will automatically appear in your spreadsheet, populated with the date, the title, and a direct link to the polished draft Google Doc!

---

## 🛠️ Advanced: Local Developer Flow (Clasp)

If you prefer building and deploying locally:
1. Copy the clasp template:
   ```bash
   cp google-apps-script/.clasp.example.json google-apps-script/.clasp.json
   ```
2. Open `google-apps-script/.clasp.json` and put the Script ID of your response sheet's Apps Script project (Script ID is in Project Settings).
3. Install clasp, log in, and push:
   ```bash
   npm install -g @google/clasp
   clasp login
   clasp push
   ```

Happy coding, and let the vibes flow! 🌊
