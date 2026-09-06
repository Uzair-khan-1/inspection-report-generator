# 📄 AI Resume Tailor

Paste your resume details and a job description — get back a tailored, fact-accurate,
ATS-friendly resume as **both Word (.docx) and PDF**, plus an honest fit analysis.
Built with **Python + Streamlit + Groq (free LLM API)**.

## What it does

1. You paste your **resume details** as plain text — your complete, real background.
   This is the only source of facts the AI is allowed to use.
2. You paste the **job description** (and any extra notes).
3. The app calls Groq's LLM to:
   - Extract required/preferred qualifications, skills, tools, and ATS keywords from
     the job description.
   - Compare them against your resume text.
   - Rewrite and reprioritize your resume content for this specific job — **without
     inventing** skills, employers, dates, certifications, or metrics.
   - Produce a **Fit Score (0–100)**, a verdict (Excellent / Strong / Moderate / Weak
     / Poor), key strengths, missing requirements, and recommendations.
4. The app populates a **fixed, built-in resume template** (`assets/resume_template.docx`)
   with the tailored content, reusing that template's own fonts/bullets/heading styles,
   and gives you both a `.docx` and a `.pdf` to download.

### Guardrails built in

- The AI is explicitly instructed to never invent skills, experience, achievements,
  certifications, employers, education, or metrics — only rephrase and reprioritize
  what's actually in the resume text you paste.
- Responses are parsed and schema-validated; malformed AI output is caught and
  surfaced as a clear error instead of silently producing a broken resume.
- Nothing is persisted — no database, no server-side storage. Everything lives in
  the Streamlit session while you use it.

## How template formatting is preserved

The app treats the built-in template's existing paragraphs as **formatting
exemplars**: it clones the underlying XML of a template's bullet point (or heading,
or job-title line) for every new bullet/line it needs to insert, so fonts, bullet
styles, spacing, and bold/italic formatting carry over automatically — even when your
tailored resume needs a different number of bullets or jobs than the template
currently has. If your pasted resume has genuinely relevant content the template has
no section for, the app adds a new heading in a matching style rather than dropping
that content.

To use a different template design, replace `assets/resume_template.docx` with your
own `.docx` (keep the section headings recognizable — e.g. "Professional Summary",
"Work Experience", "Education" — see `core/docx_utils.py`'s `SECTION_KEYWORDS` for the
full list of recognized headings, or add your own).

## Project structure

```
resume-tailor/
├── app.py                     # Streamlit UI + orchestration
├── assets/
│   └── resume_template.docx    # the fixed resume design used for every output
├── core/
│   ├── config.py               # env / st.secrets loading
│   ├── prompts.py               # system + user prompt construction
│   ├── groq_client.py           # Groq API call, JSON parsing/validation, errors
│   ├── docx_utils.py             # template analysis + tailored .docx generation
│   └── pdf_utils.py               # .docx -> .pdf conversion via LibreOffice
├── requirements.txt
├── packages.txt                # apt packages for Streamlit Cloud (LibreOffice)
├── .env.example
├── .streamlit/secrets.toml.example
├── .gitignore
└── README.md
```

## Run locally

**1. Clone and set up a virtual environment**

```bash
git clone https://github.com/<your-username>/resume-tailor.git
cd resume-tailor
python3 -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

**2. Install LibreOffice (for PDF export)**

The Word (.docx) download works without this — LibreOffice is only needed for the
PDF download button.

- **Ubuntu/Debian**: `sudo apt install libreoffice`
- **macOS**: `brew install --cask libreoffice`
- **Windows**: download from [libreoffice.org](https://www.libreoffice.org/download/)
  and make sure `soffice.exe` is on your PATH

**3. Get a free Groq API key**

Sign up at [console.groq.com](https://console.groq.com/keys) and create a key
(no credit card required).

**4. Configure your API key**

```bash
cp .env.example .env
```

Edit `.env`:

```
GROQ_API_KEY=your_actual_key_here
GROQ_MODEL=openai/gpt-oss-120b
```

> Model names change over time — check
> [console.groq.com/docs/models](https://console.groq.com/docs/models) for the
> current list of models available on the free tier if the default stops working.

**5. Run the app**

```bash
streamlit run app.py
```

Open the local URL Streamlit prints (usually `http://localhost:8501`).

## Deploy to Streamlit Community Cloud

1. Push this project to a **public or private GitHub repo** — including
   `assets/resume_template.docx` and `packages.txt`.
2. Go to [share.streamlit.io](https://share.streamlit.io) and click **New app**.
3. Pick your repo, branch, and set the main file path to `app.py`.
4. Before (or after) deploying, open **App settings → Secrets** and paste:

   ```toml
   GROQ_API_KEY = "your_actual_key_here"
   GROQ_MODEL = "openai/gpt-oss-120b"
   ```

   (This mirrors `.streamlit/secrets.toml.example` in this repo — never commit a
   real `secrets.toml` to git.)
5. Deploy. Streamlit Cloud reads `packages.txt` and installs LibreOffice via apt
   automatically before your app starts, so PDF export works out of the box — no
   extra setup needed. The app reads `GROQ_API_KEY` from `st.secrets` automatically
   on Cloud, and from `.env` locally — no code changes needed either way.

## Error handling

The app is built to fail gracefully and tell you what went wrong, rather than
crashing:

- **Empty inputs** (no resume text, no job description) → the "Analyze & Generate"
  button stays disabled with a note on what's still needed.
- **Groq API failures** (bad key, rate limit, network issue, service error) → each
  is caught individually and shown with an actionable message.
- **Malformed AI responses** (non-JSON, missing fields) → detected via schema
  validation; you're asked to retry rather than getting a broken resume.
- **Document generation failures** → caught separately so you still see your fit
  analysis even if the `.docx` build fails.
- **PDF conversion failures** (e.g. LibreOffice missing) → caught separately; the
  `.docx` download still works even if the PDF button is unavailable.

## Limitations

- This app does not use a database — each session is independent and nothing is
  saved. Refreshing the page clears your inputs.
- The tailoring quality depends on how much detail your pasted resume text actually
  contains — the AI cannot tailor what isn't there, by design.
- All output uses the one built-in template design (`assets/resume_template.docx`).
  Swap that file to change the look for everyone using this deployment.
- PDF export requires LibreOffice; on Streamlit Cloud this is handled automatically
  via `packages.txt`, but local setups need it installed separately.

## License

MIT — use, modify, and deploy freely.
