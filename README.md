# 🌟 Universal AI Newsletter Pipeline

An elegant, zero-overhead Google Apps Script automation that turns raw Google Form submissions into fully written, polished newsletter drafts using Gemini AI, and logs them directly inside the same response spreadsheet.

This repository is designed specifically for **vibecoders** and non-technical people who want a premium, friction-free "plug-and-play" pipeline. It features a **1-click automatic setup** that builds your form, links your sheets, and installs triggers in under 3 minutes!

---

## ⚡ The Flow

```
                      [ User Submits Form ]
                  (Raw Story, Details, Links,
                   Visual Assets folder link)
                               │
                               ▼
                   [ Google Sheet Response ]
                               │
                               ▼
            [ Apps Script Trigger: onFormSubmit ]
             ├── 1. Q&A Packaging (all fields → structured block)
             ├── 2. Gemini AI → Polished Article Body
             ├── 3. Gemini AI → Short Engaging Title (≤ 10 words)
             └── 4. Creates Two Google Docs in Drive
                        ├── Q&A Doc  (raw source)
                        └── Draft Doc (polished article
                             + 📸 Visual Assets link appended)
                               │
                               ▼
         [ Logs Draft Title & Link to "Drafts Log" Tab ]
```

---

## 💎 Why This Pipeline Rules

*   **1-Click Automated Setup:** Run a single helper function and Google will build the 10-question form, connect it to your sheet, create logging tabs, and register triggers programmatically. No manual clicking required!
*   **Only One Spreadsheet Needed:** No secondary tracker spreadsheets to configure. Everything is saved and logged in the same response sheet.
*   **Adaptive Gemini Prompt:** The AI reads the raw submission content, automatically analyzes the style (recognizing if it highlights a person, recaps a past event, shares a new launch, or announces a partnership), and writes the newsletter draft in the perfect style.
*   **AI-Generated Titles:** Gemini writes a short, punchy title (≤ 10 words) from the full submission context — no more pulling a raw form field as the headline. The same AI title is used for the Q&A doc, the draft doc filename, and the Drafts Log entry.
*   **Visual Assets Link in the Draft:** The contributor's submitted visual assets link (Drive, Dropbox, etc.) is appended directly to the bottom of the draft Google Doc, so the newsletter editor can access files without ever opening the original form response.

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
2. Submit a test response to the form. Include a Google Drive or Dropbox link in the **Visual Assets** field.
3. Go back to your Apps Script editor, open `FormTrigger.gs`, select `testWithLastResponse` in the top dropdown, and click **Run**.
4. A new **"Drafts Log"** tab will automatically appear in your spreadsheet, populated with the date, the title, and a direct link to the polished draft Google Doc!
5. Open the draft Doc — the AI-written article is at the top, and the **📸 Visual Assets** link is clearly labeled at the bottom for your editor.

---

## 🛠️ Advanced: Local Developer Flow (Clasp)

If you prefer building and deploying locally, see [FORM_IDS.md](google-apps-script/FORM_IDS.md) for instructions on finding your Script ID.

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
