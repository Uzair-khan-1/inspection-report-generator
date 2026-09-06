"""
AI Resume Tailor
================
Paste your master resume details and a job description; get back a resume
tailored to that job -- using a fixed, pre-designed template -- as both
.docx and .pdf, plus an honest fit analysis. Powered by Groq's free-tier LLM API.

Run locally:   streamlit run app.py
Deploy:        Streamlit Community Cloud (see README.md)
"""

import streamlit as st

from core.config import get_bundled_template_path, get_groq_api_key, get_groq_model
from core.docx_utils import detect_template_sections, generate_tailored_docx
from core.groq_client import ResumeTailorError, analyze_and_tailor
from core.pdf_utils import PdfConversionError, convert_docx_to_pdf, is_libreoffice_available

st.set_page_config(page_title="AI Resume Tailor", page_icon="📄", layout="centered")

VERDICT_COLORS = {
    "Excellent": "🟢",
    "Strong": "🟢",
    "Moderate": "🟡",
    "Weak": "🟠",
    "Poor": "🔴",
}


@st.cache_data
def _load_template_bytes() -> bytes:
    with open(get_bundled_template_path(), "rb") as f:
        return f.read()


@st.cache_data
def _load_template_sections() -> list[str]:
    return detect_template_sections(_load_template_bytes())


def _init_state():
    defaults = {
        "master_resume_text": "",
        "job_description": "",
        "job_details": "",
        "result": None,        # parsed AI response dict
        "output_docx": None,   # generated tailored resume bytes
        "output_pdf": None,    # generated tailored resume PDF bytes
        "pdf_error": None,
        "error": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def main():
    _init_state()

    st.title("📄 AI Resume Tailor")
    st.caption(
        "Paste your resume details and a job description, and get a tailored, "
        "ATS-friendly resume -- as Word and PDF -- with an honest fit analysis."
    )

    api_key = get_groq_api_key()
    model = get_groq_model()
    if not api_key:
        st.warning(
            "⚠️ No Groq API key found. Add `GROQ_API_KEY` to a local `.env` file, "
            "or to your Streamlit Cloud app's **Secrets**. See README.md for details.",
            icon="⚠️",
        )
    if not is_libreoffice_available():
        st.info(
            "ℹ️ LibreOffice isn't detected in this environment, so PDF export will be "
            "unavailable this session -- the Word (.docx) download will still work. "
            "See README.md to enable PDF export.",
            icon="ℹ️",
        )

    try:
        sections = _load_template_sections()
        st.caption("Using the built-in resume template. Sections it will populate: " + " • ".join(sections))
    except Exception as e:  # noqa: BLE001
        st.error(f"Couldn't load the built-in resume template: {e}")
        st.stop()

    # ---- Step 1: Master resume ----
    st.header("1. Paste your resume details")
    st.session_state.master_resume_text = st.text_area(
        "Your complete, real background -- everything you've ever done. The more "
        "complete this is, the better the AI can tailor your resume without inventing anything.",
        value=st.session_state.master_resume_text,
        height=300,
        placeholder=(
            "Paste your full resume content here: name, contact info, summary, skills, "
            "work experience with dates and bullet points, education, certifications, "
            "training, languages, etc."
        ),
    )

    # ---- Step 2: Job description ----
    st.header("2. Paste the job description")
    st.session_state.job_description = st.text_area(
        "Job description (required)",
        value=st.session_state.job_description,
        height=220,
        placeholder="Paste the full job posting here...",
    )
    st.session_state.job_details = st.text_area(
        "Any other job details (optional)",
        value=st.session_state.job_details,
        height=100,
        placeholder="E.g. company name, seniority level, must-have tools not mentioned above, notes from a recruiter call...",
    )

    # ---- Step 3: Analyze & Generate ----
    st.header("3. Analyze & Generate")

    ready = bool(st.session_state.master_resume_text.strip() and st.session_state.job_description.strip())
    missing = []
    if not st.session_state.master_resume_text.strip():
        missing.append("your resume details")
    if not st.session_state.job_description.strip():
        missing.append("the job description")
    if missing and not ready:
        st.info("Still needed: " + ", ".join(missing) + ".")

    generate_clicked = st.button("✨ Analyze & Generate", type="primary", disabled=not ready)

    if generate_clicked:
        st.session_state.result = None
        st.session_state.output_docx = None
        st.session_state.output_pdf = None
        st.session_state.pdf_error = None
        st.session_state.error = None

        with st.spinner("Analyzing the job description and tailoring your resume... this can take up to a minute."):
            try:
                data = analyze_and_tailor(
                    master_resume_text=st.session_state.master_resume_text,
                    job_description=st.session_state.job_description,
                    job_details=st.session_state.job_details,
                    api_key=api_key,
                    model=model,
                )
                st.session_state.result = data
                try:
                    docx_bytes = generate_tailored_docx(_load_template_bytes(), data["tailored_resume"])
                    st.session_state.output_docx = docx_bytes
                    try:
                        st.session_state.output_pdf = convert_docx_to_pdf(docx_bytes)
                    except PdfConversionError as e:
                        st.session_state.pdf_error = str(e)
                except Exception as e:  # noqa: BLE001
                    st.session_state.error = (
                        "The AI analysis succeeded, but building the final .docx failed: "
                        f"{e}. Your fit analysis below is still valid -- try generating again."
                    )
            except ResumeTailorError as e:
                st.session_state.error = str(e)
            except Exception as e:  # noqa: BLE001
                st.session_state.error = f"Unexpected error: {e}"

    if st.session_state.error:
        st.error(st.session_state.error)

    # ---- Results ----
    if st.session_state.result:
        fit = st.session_state.result.get("fit_analysis", {})
        st.header("Fit Analysis")

        score = fit.get("fit_score", "?")
        verdict = fit.get("verdict", "Unknown")
        icon = VERDICT_COLORS.get(verdict, "⚪")

        col1, col2 = st.columns([1, 2])
        with col1:
            st.metric("Fit Score", f"{score}/100")
        with col2:
            st.markdown(f"### {icon} {verdict}")

        strengths = fit.get("key_strengths", [])
        missing_reqs = fit.get("missing_requirements", [])
        recs = fit.get("recommendations", [])

        if strengths:
            st.subheader("✅ Key Strengths")
            for s in strengths:
                st.markdown(f"- {s}")

        if missing_reqs:
            st.subheader("⚠️ Missing / Weak Requirements")
            for m in missing_reqs:
                st.markdown(f"- {m}")

        if recs:
            st.subheader("💡 Recommendations")
            for r in recs:
                st.markdown(f"- {r}")

        st.header("Download Your Tailored Resume")
        if st.session_state.output_docx:
            name_part = (st.session_state.result.get("tailored_resume", {}) or {}).get("name", "").strip()
            base_name = "".join(c for c in name_part if c.isalnum() or c in " _-").strip().replace(" ", "_") or "tailored_resume"

            dcol, pcol = st.columns(2)
            with dcol:
                st.download_button(
                    "⬇️ Download Word (.docx)",
                    data=st.session_state.output_docx,
                    file_name=f"{base_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True,
                )
            with pcol:
                if st.session_state.output_pdf:
                    st.download_button(
                        "⬇️ Download PDF",
                        data=st.session_state.output_pdf,
                        file_name=f"{base_name}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )
                else:
                    st.button("PDF unavailable", disabled=True, use_container_width=True)
                    if st.session_state.pdf_error:
                        st.caption(st.session_state.pdf_error)

            st.caption(
                "Tip: open the file and do a quick read-through -- the AI never invents "
                "facts, but always proofread before sending."
            )
        else:
            st.info("Resume document isn't available for this run (see error above).")

    st.divider()
    with st.expander("About / How this works"):
        st.markdown(
            """
- **Nothing is invented.** The AI is instructed to only use facts, employers, dates,
  skills, and metrics that already exist in the resume text you paste in. It rewrites
  and reprioritizes -- it doesn't fabricate.
- **One fixed template.** This app always outputs using the same built-in resume
  design, so formatting stays consistent across every job you tailor for.
- **Nothing is stored.** This app has no database; your resume and job description
  only exist in memory for this session.
            """
        )


if __name__ == "__main__":
    main()
