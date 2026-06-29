# 🗞️ PressKit
### An AI-Powered Newsletter Pipeline

Collect community stories with a Google Form → auto-generate polished newsletter drafts with Gemini AI → log everything directly inside the same Google Sheet. Zero overhead, zero manual writing.

Built for **vibe coders and non-technical teams** who want a premium, plug-and-play content pipeline. One paste, one run — your entire form, sheet, trigger, and Drafts Log are set up automatically in under 3 minutes.

---

## ⚡ How It Works

```
              [ Community Member Submits Form ]
           (Name, Role, Story, Links, Visual Assets)
                              │
                              ▼
                  [ Google Sheet Response ]
                              │
                              ▼
           [ Apps Script Trigger: onFormSubmit ]
            ├── 1. Q&A Packaging (fields → structured block)
            ├── 2. Gemini AI → Polished Article Body
            ├── 3. Gemini AI → Short Engaging Title (≤ 10 words)
            └── 4. Creates Two Google Docs in Drive
                       ├── Q&A Doc  (raw source reference)
                       └── Draft Doc (polished article
                            + 📸 Visual Assets link appended)
                              │
                              ▼
       [ Logs Title & Link to "Drafts Log" Sheet ]
```

---

## 💎 Why PressKit

- **1-Click Setup:** `setupPipeline()` builds the 10-question form, links it to your sheet, creates the Drafts Log, and registers the trigger. No manual clicking.
- **Everything in One Sheet:** Drafts Log lives alongside form responses — no separate tracker to manage.
- **Adaptive AI Prompt:** Gemini reads the submission and automatically writes in the right style — member spotlight, event recap, partnership announcement, or launch post.
- **AI-Generated Titles:** The article headline comes from the finished draft, not a raw form field. Same title is used for the Doc filename, Q&A doc, and Drafts Log entry.
- **Visual Assets in Every Draft:** The contributor's Drive/Dropbox link is appended directly to the draft Doc so editors never have to dig through the form response.
- **Failure-Proof:** If Gemini is unavailable, the row is logged as `⚠️ FAILED` instead of creating a broken draft. Retry failed rows in one click from the **📋 PressKit** spreadsheet menu.
- **AI Test Data:** Generate and submit realistic test form responses on demand — Gemini creates 3 diverse submissions (spotlight, recap, announcement) and fires the full pipeline for each, with graceful retry on high-demand spikes.

---

## 🚀 Quick Start — One Paste, One Run

### 1. Paste the Bundle
1. Create a new **Google Sheet**.
2. Go to **Extensions → Apps Script**.
3. Click into the default `Code.gs` file, **select all**, and replace it with the contents of **[bundle.gs](google-apps-script/bundle.gs)**.
4. Click **Save** (💾).

### 2. Run the Initializer
1. In the function dropdown, select **`setupPipeline`** and click **Run**.
2. **Authorize Permissions** when prompted (*Advanced → Go to Untitled Project (unsafe)*).
3. Open the **Execution Log** — it will print links to edit and share your new Google Form.

### 3. Add Your Gemini API Key
1. Get a free key at [Google AI Studio](https://aistudio.google.com/).
2. In Apps Script → **Project Settings (⚙️)** → **Script Properties** → **Add script property**.
3. Name: `GEMINI_API_KEY` · Value: your key. Save.

> **Optional:** Set `ORG_NAME` in the `Config` object at the top of the bundle to tailor AI content to your organization. Leave it blank (`''`) for generic output.

> **Optional:** Set `DRAFT_FOLDER_ID` to save draft Docs to a specific Drive folder instead of the spreadsheet's parent folder.

---

## 🎨 Test the Pipeline

After setup, open the **📋 PressKit** menu in your spreadsheet toolbar:

| Menu Item | What It Does |
|---|---|
| **▶ Run Pipeline on Last Response** | Runs the full AI pipeline on the most recent form row |
| **🔄 Reprocess Failed Drafts** | Retries any rows logged as `⚠️ FAILED` |
| **🧪 Submit Test Form Data** | Asks Gemini to generate 3 diverse test submissions and fires the full pipeline for each |

---

## 🛠️ Multi-File Setup (Optional)

Prefer organized separate files? Create each of the following in Apps Script (click **+** next to Files) and paste the matching file:

- [Config.gs](google-apps-script/Config.gs)
- [ArticleBuilder.gs](google-apps-script/ArticleBuilder.gs)
- [GeminiArticle.gs](google-apps-script/GeminiArticle.gs)
- [appendArticleTask.gs](google-apps-script/appendArticleTask.gs)
- [FormTrigger.gs](google-apps-script/FormTrigger.gs)
- [FormSetup.gs](google-apps-script/FormSetup.gs)

Then follow Steps 2 and 3 — `setupPipeline` and the API key setup are identical.

---

## 🧑‍💻 Developer Setup (Clasp)

To build and deploy locally, see [FORM_IDS.md](google-apps-script/FORM_IDS.md) for Script ID instructions.

```bash
cp google-apps-script/.clasp.example.json google-apps-script/.clasp.json
# Fill in your Script ID in .clasp.json, then:
npm install -g @google/clasp
clasp login
clasp push
```

---

*PressKit — AI Newsletter Pipeline. Built for the community, by @adigitaltati.* 🌊
