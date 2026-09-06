"""
DOCX -> PDF conversion.

Streamlit has no built-in PDF export, and Word formatting (fonts, bullets,
spacing) is genuinely hard to reproduce faithfully with a pure-Python PDF
library. The reliable, widely-used approach is headless LibreOffice, which
renders the .docx exactly as it would print and converts it losslessly.

Local dev: install LibreOffice (e.g. `sudo apt install libreoffice` on
Ubuntu/Debian, `brew install --cask libreoffice` on macOS, or download from
libreoffice.org on Windows) and make sure `soffice` is on your PATH.

Streamlit Community Cloud: add a `packages.txt` file (already included in
this project) with a `libreoffice` line -- Streamlit Cloud installs it via
apt automatically before your app starts.
"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path


class PdfConversionError(Exception):
    """Raised when LibreOffice isn't available or the conversion fails."""


def is_libreoffice_available() -> bool:
    return shutil.which("soffice") is not None or shutil.which("libreoffice") is not None


def _soffice_binary() -> str:
    return shutil.which("soffice") or shutil.which("libreoffice") or "soffice"


def convert_docx_to_pdf(docx_bytes: bytes, timeout_seconds: int = 60) -> bytes:
    """Convert .docx bytes to .pdf bytes via headless LibreOffice.

    Raises PdfConversionError with a friendly message on any failure so the
    caller can show it without the app crashing -- the .docx download should
    still work even if PDF conversion isn't available in this environment.
    """
    if not is_libreoffice_available():
        raise PdfConversionError(
            "PDF export isn't available in this environment because LibreOffice "
            "isn't installed. The Word (.docx) download still works. To enable "
            "PDF export: locally, install LibreOffice; on Streamlit Cloud, make "
            "sure packages.txt (containing 'libreoffice') is in your repo and "
            "redeploy."
        )

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        docx_path = tmp_path / "resume.docx"
        docx_path.write_bytes(docx_bytes)

        try:
            result = subprocess.run(
                [
                    _soffice_binary(),
                    "--headless",
                    "--norestore",
                    "--convert-to", "pdf",
                    "--outdir", str(tmp_path),
                    str(docx_path),
                ],
                capture_output=True,
                timeout=timeout_seconds,
                check=False,
            )
        except subprocess.TimeoutExpired as e:
            raise PdfConversionError(
                "PDF conversion timed out. Please try again -- the .docx "
                "download is still available."
            ) from e
        except Exception as e:  # noqa: BLE001
            raise PdfConversionError(f"PDF conversion failed to start: {e}") from e

        pdf_path = tmp_path / "resume.pdf"
        if result.returncode != 0 or not pdf_path.exists():
            stderr = (result.stderr or b"").decode(errors="ignore")[:500]
            raise PdfConversionError(
                "PDF conversion failed. The .docx download is still available. "
                f"Details: {stderr or 'unknown LibreOffice error'}"
            )

        return pdf_path.read_bytes()
