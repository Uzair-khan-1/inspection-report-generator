"""Thin wrapper around the Groq chat completions API for resume analysis + tailoring."""

from __future__ import annotations

import json
import re

from groq import Groq
from groq import (
    APIConnectionError,
    APIStatusError,
    AuthenticationError,
    RateLimitError,
)

from .prompts import SYSTEM_PROMPT, build_user_prompt

REQUIRED_FIT_KEYS = {"fit_score", "verdict", "key_strengths", "missing_requirements", "recommendations"}
REQUIRED_RESUME_KEYS = {
    "name", "title", "contact", "summary", "skills", "experience",
    "education", "certifications", "training", "projects", "languages", "additional_sections",
}


class ResumeTailorError(Exception):
    """User-facing error with a friendly message; raised for any failure in
    the AI call or response parsing so app.py can show a clean st.error()."""


def _strip_code_fences(text: str) -> str:
    text = text.strip()
    fence_match = re.match(r"^```(?:json)?\s*(.*?)\s*```$", text, re.DOTALL)
    if fence_match:
        return fence_match.group(1).strip()
    return text


def _extract_json_object(text: str) -> str:
    """As a fallback, grab the first {...} block in the text in case the
    model added stray commentary despite instructions."""
    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        return text[start:end + 1]
    return text


def _parse_ai_json(raw_text: str) -> dict:
    candidate = _strip_code_fences(raw_text)
    try:
        return json.loads(candidate)
    except json.JSONDecodeError:
        pass

    candidate2 = _extract_json_object(candidate)
    try:
        return json.loads(candidate2)
    except json.JSONDecodeError as e:
        raise ResumeTailorError(
            "The AI returned a response that wasn't valid JSON, so it can't be used "
            "to build your resume. This sometimes happens on a busy request -- "
            "please try clicking 'Analyze & Generate' again."
        ) from e


def _validate_schema(data: dict) -> None:
    if "fit_analysis" not in data or "tailored_resume" not in data:
        raise ResumeTailorError(
            "The AI response was missing required sections. Please try again."
        )
    fit = data["fit_analysis"]
    resume = data["tailored_resume"]
    if not isinstance(fit, dict) or not REQUIRED_FIT_KEYS.issubset(fit.keys()):
        raise ResumeTailorError(
            "The AI's fit analysis was incomplete. Please try again."
        )
    if not isinstance(resume, dict) or not REQUIRED_RESUME_KEYS.issubset(resume.keys()):
        raise ResumeTailorError(
            "The AI's tailored resume content was incomplete. Please try again."
        )


def analyze_and_tailor(
    master_resume_text: str,
    job_description: str,
    job_details: str,
    api_key: str,
    model: str,
) -> dict:
    """Call Groq to analyze fit and produce tailored resume content.

    Returns a dict with keys 'fit_analysis' and 'tailored_resume'.
    Raises ResumeTailorError with a user-friendly message on any failure.
    """
    if not api_key:
        raise ResumeTailorError(
            "No Groq API key configured. Add GROQ_API_KEY to your .env file "
            "(local) or to Streamlit Cloud secrets (deployed)."
        )
    if not master_resume_text or not master_resume_text.strip():
        raise ResumeTailorError("Your master resume text is empty -- please upload a file or paste your resume.")
    if not job_description or not job_description.strip():
        raise ResumeTailorError("The job description is empty -- please paste the job description.")

    client = Groq(api_key=api_key)
    user_prompt = build_user_prompt(master_resume_text, job_description, job_details)

    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_prompt},
    ]

    def _call(with_json_mode: bool):
        kwargs = dict(
            model=model,
            messages=messages,
            temperature=0.4,
            max_tokens=4096,
        )
        if with_json_mode:
            kwargs["response_format"] = {"type": "json_object"}
        return client.chat.completions.create(**kwargs)

    try:
        try:
            completion = _call(with_json_mode=True)
        except APIStatusError:
            # Some Groq models reject response_format; retry in plain mode.
            completion = _call(with_json_mode=False)
    except AuthenticationError as e:
        raise ResumeTailorError(
            "Groq rejected the API key. Double-check GROQ_API_KEY in your .env "
            "or Streamlit secrets."
        ) from e
    except RateLimitError as e:
        raise ResumeTailorError(
            "Groq's rate limit was hit. Wait a moment and try again, or check "
            "your Groq account's usage limits."
        ) from e
    except APIConnectionError as e:
        raise ResumeTailorError(
            "Couldn't reach the Groq API. Check your internet connection and try again."
        ) from e
    except APIStatusError as e:
        raise ResumeTailorError(
            f"Groq API returned an error (status {getattr(e, 'status_code', '?')}). "
            "Please try again in a moment."
        ) from e
    except Exception as e:  # noqa: BLE001 - last-resort catch for a clean UI message
        raise ResumeTailorError(f"Unexpected error calling the AI service: {e}") from e

    if not completion.choices:
        raise ResumeTailorError("The AI returned an empty response. Please try again.")

    raw_text = completion.choices[0].message.content or ""
    data = _parse_ai_json(raw_text)
    _validate_schema(data)
    return data
