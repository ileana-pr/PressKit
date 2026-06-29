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

## 🚀 Quick Start (One Paste, One Run)

### 1. Paste the Bundle
1. Create a brand new **Google Sheet**.
2. Go to the top menu and select **Extensions ➔ Apps Script**.
3. Click into the default `Code.gs` file, **select all**, and replace it with the contents of **[bundle.gs](google-apps-script/bundle.gs)**.
4. Click **Save** (💾).

### 2. Run the Initializer
1. In the function dropdown at the top, select **`setupPipeline`** and click **Run**.
2. **Authorize Permissions:** Follow the prompt (click *Advanced* ➔ *Go to Untitled Project (unsafe)*).
3. Once it finishes, open the **Execution Log** — it will print the links to edit and share your new Google Form!

### 3. Add Your Gemini API Key
1. Get a free key from [Google AI Studio](https://aistudio.google.com/).
2. In Apps Script, click **Project Settings (gear icon)** ➔ **Script Properties** ➔ **Add script property**.
3. Set the property name to `GEMINI_API_KEY` and paste your key as the value. Save.

> **Optional:** To save draft Docs to a specific Drive folder, edit the `DRAFT_FOLDER_ID` value near the top of the bundle (find the ID in your folder's URL: `drive.google.com/drive/folders/YOUR_ID_HERE`).

---

### 🎨 The Test Run (Watch the Magic!)
1. Open the shareable Google Form link you copied in Step 2.
2. Submit a test response. Include a Google Drive or Dropbox link in the **Visual Assets** field.
3. Back in the Apps Script editor, select **`testWithLastResponse`** in the function dropdown and click **Run**.
4. The **Drafts Log** sheet will populate with the date, AI-generated title, and a direct link to the polished draft Doc!
5. Open the draft Doc — the AI-written article is at the top, and the **📸 Visual Assets** link is clearly labeled at the bottom for your editor.

---

## 🛠️ Alternative: Multi-File Setup

Prefer to keep the code organized across separate files? Create each of the following in the Apps Script editor (click **+** next to Files) and paste the matching file from this repo:

*   [Config.gs](google-apps-script/Config.gs)
*   [ArticleBuilder.gs](google-apps-script/ArticleBuilder.gs)
*   [GeminiArticle.gs](google-apps-script/GeminiArticle.gs)
*   [appendArticleTask.gs](google-apps-script/appendArticleTask.gs)
*   [FormTrigger.gs](google-apps-script/FormTrigger.gs)
*   [FormSetup.gs](google-apps-script/FormSetup.gs)

Then follow Steps 2 and 3 above — `setupPipeline` and the API key setup are the same.

---

## 🧑‍💻 Developer Flow (Clasp)

If you prefer to build and deploy locally, see [FORM_IDS.md](google-apps-script/FORM_IDS.md) for instructions on finding your Script ID.

1. Copy the clasp template:
   ```bash
   cp google-apps-script/.clasp.example.json google-apps-script/.clasp.json
   ```
2. Open `google-apps-script/.clasp.json` and fill in the Script ID of your response sheet's Apps Script project.
3. Install clasp, log in, and push:
   ```bash
   npm install -g @google/clasp
   clasp login
   clasp push
   ```

Happy coding, and let the vibes flow! 🌊
