# Script IDs (for Clasp Developer Setup)

With the **Consolidated Form** update, you no longer need to manage and deploy to multiple separate form response sheet script IDs! There is now only **one script project** attached to your single consolidated form response sheet.

---

## 🛠️ How to Find Your Script ID

**IMPORTANT:** Always use the **Apps Script Project ID (Script ID)** in clasp, *not* the Google Sheet spreadsheet ID.

1. Open your **consolidated form response spreadsheet** in your browser.
2. Go to the top menu and select **Extensions ➔ Apps Script**.
3. In the left-hand sidebar of the Apps Script editor, click the **Project Settings (gear icon)**.
4. Copy the **Script ID** (under the "IDs" section).
5. Open your local project, go to `google-apps-script/.clasp.json`, and replace the `scriptId` value with your copied ID:
   ```json
   {
     "scriptId": "YOUR_COPIED_SCRIPT_ID"
   }
   ```
6. Now you can run `clasp push` from the root of the repository to deploy all script files directly into the sheet's Apps Script project!

---

## 🔄 Automated Universal Mapping

The pipeline now automatically maps and formats all incoming community submissions, generating Google Doc drafts and logging them directly inside the same response spreadsheet:

| Submission Field | Script Action | Logged Tab Location |
| :--- | :--- | :--- |
| **Name / Summary Column** | Auto-identified and truncated to neat length | `'Drafts Log'` tab (`Date`, `Draft Title`, `Google Doc Link`) |
| **All Generic Q&A Fields** | Formatted into a clean structured block for Gemini | Logged as deliverables in spreadsheet |
