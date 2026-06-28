# Script IDs (for Clasp Developer Setup)

There is only **one script project** in this pipeline — attached to the single consolidated form response spreadsheet.

---

## 🛠️ How to Find Your Script ID

**IMPORTANT:** Always use the **Apps Script Project ID (Script ID)** in clasp, *not* the Google Sheet spreadsheet ID.

1. Open your **form response spreadsheet** in your browser.
2. Go to the top menu and select **Extensions ➔ Apps Script**.
3. In the left-hand sidebar of the Apps Script editor, click the **Project Settings (gear icon)**.
4. Copy the **Script ID** (under the "IDs" section).
5. Open your local project, go to `google-apps-script/.clasp.json`, and replace the `scriptId` value with your copied ID:
   ```json
   {
     "scriptId": "YOUR_COPIED_SCRIPT_ID"
   }
   ```
6. Now you can run `clasp push` from the root of the repository to deploy all script files directly into the sheet's Apps Script project.

---

## 📋 How Form Fields Map Through the Pipeline

| Form Field | Script Action | Where It Ends Up |
| :--- | :--- | :--- |
| **All Q&A Fields** | Packaged into a structured Q/A block | Q&A Doc (raw source); fed to Gemini for article + title |
| **Gemini-generated title** | Produced from full submission context (≤ 10 words) | Google Doc filenames + `Drafts Log` tab (`Draft Title` column) |
| **Visual Assets** (Drive/Dropbox link) | Extracted separately; not rewritten by AI | Appended as **📸 Visual Assets** section at the bottom of the draft Doc |
| **Consent & Attribution** | Captured in Q&A block; Gemini is instructed to exclude it from the article | Q&A Doc only |

---

The `Drafts Log` tab always contains: `Date`, `Draft Title`, and `Google Doc Link` (pointing to the polished AI draft).
