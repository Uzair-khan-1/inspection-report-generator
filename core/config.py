"""
Configuration loader.

Works both locally (via a .env file + python-dotenv) and on Streamlit
Community Cloud (via st.secrets). Local .env always takes priority if
present, so you can override secrets during development.
"""

import os

from dotenv import load_dotenv

load_dotenv()  # no-op if .env doesn't exist


def get_setting(key: str, default: str | None = None) -> str | None:
    """
    Fetch a setting by name, checking (in order):
      1. Environment variables / .env file
      2. Streamlit secrets (st.secrets), if running inside Streamlit
      3. The provided default
    """
    value = os.environ.get(key)
    if value:
        return value

    try:
        import streamlit as st

        if hasattr(st, "secrets") and key in st.secrets:
            return st.secrets[key]
    except Exception:
        # st.secrets raises if no secrets.toml exists at all; ignore.
        pass

    return default


def get_groq_api_key() -> str | None:
    return get_setting("GROQ_API_KEY")


def get_groq_model() -> str:
    return get_setting("GROQ_MODEL", "openai/gpt-oss-120b")


def get_bundled_template_path() -> str:
    """Path to the fixed resume template shipped with this app (users no
    longer upload their own template -- see assets/resume_template.docx)."""
    import os

    return os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets", "resume_template.docx")
