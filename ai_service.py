"""
ai_service.py
--------------
Turns rough user text into short, professional report text using Groq's
free-tier API (https://console.groq.com) with "llama-3.1-8b-instant" -- one
of the fastest hosted models available (sub-second responses), which keeps
the app faster than typing the same thing by hand in Excel.

Design goals (per spec):
  * Text only -- photos are NEVER sent to the AI.
  * ONE request rewrites every pending item's Location + Finding at once,
    to minimize latency and API calls.
  * If the AI call fails for any reason (no key, network, bad JSON, rate
    limit, timeout...) we silently fall back to the user's original text
    (lightly cleaned up) so the app never blocks the user.
  * The model is instructed not to invent facts, only rephrase.
"""

import json
import os
import re

import requests

GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_MODEL = "llama-3.1-8b-instant"
REQUEST_TIMEOUT = 12  # seconds -- fail fast and fall back rather than hang

SYSTEM_PROMPT = (
    "You are an assistant that rewrites rough facility-inspection notes into "
    "short, professional report language. Rules:\n"
    "1. Do NOT invent, assume, or add any information that is not present in "
    "the input.\n"
    "2. 'location' must become a clean 3-4 word Title Case location/building "
    "name (no extra commentary).\n"
    "3. 'finding' must become one short, professional sentence describing the "
    "same issue, ending with a period.\n"
    "4. Return ONLY a JSON array, no prose, no markdown fences, one object per "
    "input item in the same order, each shaped exactly like: "
    '{"location": "...", "finding": "..."}'
)


def _get_api_key():
    # Works both in Streamlit Cloud (st.secrets) and locally (env var)
    try:
        import streamlit as st
        if "GROQ_API_KEY" in st.secrets:
            return st.secrets["GROQ_API_KEY"]
    except Exception:
        pass
    return os.environ.get("GROQ_API_KEY")


def _fallback_clean(raw_location, raw_finding):
    """Best-effort local cleanup used only if the AI call fails."""
    location = " ".join(raw_location.strip().split())
    location = location.title() if location else location

    finding = " ".join(raw_finding.strip().split())
    if finding:
        finding = finding[0].upper() + finding[1:]
        if not finding.endswith((".", "!", "?")):
            finding += "."
    return location, finding


def _fallback_all(items):
    out = []
    for it in items:
        loc, find = _fallback_clean(it.get("location", ""), it.get("finding", ""))
        out.append({"location": loc, "finding": find})
    return out


def _extract_json_array(text):
    """The model should return pure JSON, but strip code fences defensively."""
    text = text.strip()
    text = re.sub(r"^```(json)?", "", text).strip()
    text = re.sub(r"```$", "", text).strip()
    return json.loads(text)


def process_items_with_ai(items):
    """
    items: list of {"location": raw_text, "finding": raw_text}
    returns: list of {"location": clean_text, "finding": clean_text}, same
             order/length as input. Falls back to local cleanup on any error.
    """
    if not items:
        return []

    api_key = _get_api_key()
    if not api_key:
        return _fallback_all(items)

    user_payload = json.dumps(
        [{"location": it.get("location", ""), "finding": it.get("finding", "")} for it in items],
        ensure_ascii=False,
    )

    try:
        resp = requests.post(
            GROQ_API_URL,
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            json={
                "model": GROQ_MODEL,
                "temperature": 0.2,
                "max_tokens": 1200,
                "messages": [
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": user_payload},
                ],
            },
            timeout=REQUEST_TIMEOUT,
        )
        resp.raise_for_status()
        content = resp.json()["choices"][0]["message"]["content"]
        result = _extract_json_array(content)

        if not isinstance(result, list) or len(result) != len(items):
            raise ValueError("AI response shape mismatch")

        cleaned = []
        for original, r in zip(items, result):
            loc = str(r.get("location", "")).strip()
            find = str(r.get("finding", "")).strip()
            if not loc or not find:
                loc_fb, find_fb = _fallback_clean(original.get("location", ""), original.get("finding", ""))
                loc = loc or loc_fb
                find = find or find_fb
            cleaned.append({"location": loc, "finding": find})
        return cleaned

    except Exception:
        # Network error, timeout, bad JSON, rate limit, missing key, etc.
        # Never block the user -- just use their original text.
        return _fallback_all(items)
