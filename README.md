# 📋 AI Inspection Report Generator

A very fast, no-frills Streamlit app that turns rough field notes + two
photos per finding into a properly formatted Excel inspection report —
using your exact company template.

**Workflow:** Location → Finding → Before photo → After photo → **Add** →
(repeat) → **Generate Report** → **Download Excel**.

No manual Excel editing, resizing, or repositioning required.

---

## How it works

| Step | What happens |
|---|---|
| You type a rough location + finding, and upload a Before/After photo | Stored locally in the browser session (nothing saved to disk/DB) |
| Click **Add item** | Repeat for as many inspection points as you like |
| Click **Generate Report** | **One** AI request cleans up *all* the text at once (fast, cheap) |
| — | Photos are resized **locally** (Pillow) and placed into the correct cells, preserving aspect ratio |
| — | The exact template (`assets/Template.xlsx`) is filled in with `openpyxl`, formatting fully preserved |
| Click **Download Excel** | Get your finished `.xlsx` report |

### Excel column mapping
| Column | Content |
|---|---|
| C | Location / Building (AI-cleaned) |
| D | Findings (AI-cleaned text) + **Before** photo |
| E | Corrective Action → **After** photo |
| F | Completed On (today's date) |
| G | Status ("Open" by default — color-codes automatically via the template's built-in conditional formatting) |

The template ships with 3 ready-made rows. If you add a 4th+ item, the app
inserts a new row block and copies the exact same fonts/borders/fills/row
height, so formatting always matches — no matter how many items you add.

## AI model

Uses **Groq's free-tier API** with `llama-3.1-8b-instant` — one of the
fastest hosted LLM endpoints available, so the AI step adds well under a
second per batch. Only the rough **text** (never photos) is sent, and only
**one request** covers every pending item.

- Get a free key at https://console.groq.com/keys
- If no key is set, or the request fails for any reason, the app
  automatically falls back to lightly-cleaned versions of your original
  text and keeps working — it never blocks you.

## Setup

```bash
pip install -r requirements.txt
```

Add your Groq key (either works):

```bash
# Option A: environment variable
export GROQ_API_KEY="your-key-here"

# Option B: Streamlit secrets (create .streamlit/secrets.toml)
echo 'GROQ_API_KEY = "your-key-here"' > .streamlit/secrets.toml
```

Run locally:

```bash
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Push this folder to a GitHub repo (the `assets/Template.xlsx` file must
   be included — it's the template the app fills in).
2. On https://share.streamlit.io, create a new app pointing at `app.py`.
3. In the app's **Settings → Secrets**, add:
   ```toml
   GROQ_API_KEY = "your-key-here"
   ```
4. Deploy. That's it.

## Project structure

```
app.py               # Streamlit UI + workflow
excel_processor.py   # Fills the template: text, photo resize/placement, row insertion
ai_service.py         # One batched Groq API call + safe local fallback
assets/Template.xlsx  # Your exact report template (do not rename/move)
requirements.txt
README.md
.gitignore
```

## Notes / known limitations

- Excel **data-validation dropdown lists** in the template (defined via a
  newer Excel extension format) are dropped when the file is re-saved by
  `openpyxl` — this is a library limitation, not a bug in the app. All
  visible **formatting, colors, and conditional formatting still work
  perfectly** (see the Status column, which stays red/green automatically).
- Designed to stay fast and simple on purpose: no database, no OCR/CV, no
  multi-page previews. If AI cleanup fails, your original text is used as-is.
