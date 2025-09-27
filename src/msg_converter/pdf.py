# src/msg_converter/pdf.py
from __future__ import annotations
from pathlib import Path

from .converters.msg_to_pdf import convert_to_pdf as _convert_to_pdf  # your implementation

class MsgToPdfError(Exception):
    """Normalized error for msg→PDF conversion failures."""

def msg_to_pdf(src: str | Path, dest: str | Path) -> None:
    """
    Public API: Convert a single .msg file to a PDF.
    - Ensures parent directories exist
    - Normalizes exceptions to MsgToPdfError
    """
    src_p = Path(src)
    dest_p = Path(dest)
    try:
        dest_p.parent.mkdir(parents=True, exist_ok=True)
        _convert_to_pdf(str(src_p), str(dest_p))
    except Exception as e:
        raise MsgToPdfError(f"{src_p} → {dest_p}: {e}") from e
