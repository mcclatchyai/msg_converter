# msg_to_csv.py
from __future__ import annotations
import csv
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import List, Tuple
from datetime import datetime
import extract_msg  # uses your existing dependency

# ---- Data model -------------------------------------------------------------
@dataclass
class EmailRow:
    file_name: str
    subject: str
    from_name: str
    from_email: str
    to: str
    cc: str
    bcc: str
    date_utc: str
    attachments_count: int
    attachments: str
    body_text: str

# ---- Helpers ----------------------------------------------------------------
def _safe(v) -> str:
    return "" if v is None else str(v).strip()

def _join(items) -> str:
    return ", ".join([_safe(i) for i in items if _safe(i)])

def _normalize_date(dt_str: str) -> str:
    """
    Try to standardize to ISO 8601 (UTC if tz present). If parsing fails, return the original.
    extract-msg often yields RFC2822-like strings (e.g., 'Thu, 01 Dec 2016 11:44:10 -0500').
    """
    if not dt_str:
        return ""
    for fmt in (
        "%a, %d %b %Y %H:%M:%S %z",
        "%d %b %Y %H:%M:%S %z",
        "%m/%d/%Y %H:%M:%S %p",
        "%Y-%m-%d %H:%M:%S%z",
        "%Y-%m-%d %H:%M:%S",
    ):
        try:
            dt = datetime.strptime(dt_str, fmt)
            return dt.astimezone().isoformat()
        except Exception:
            continue
    return dt_str  # fall back

def _recipient_list(msg, field: str) -> str:
    """
    extract_msg exposes both aggregated strings (msg.to, msg.cc, msg.bcc)
    and structured recipients on newer versions. Prefer structured if present.
    """
    try:
        recips = getattr(msg, "recipients", None) or []
        if recips:
            vals = []
            for r in recips:
                # r.type: 'To', 'Cc', 'Bcc'
                if getattr(r, "type", "").lower() == field:
                    nm = _safe(getattr(r, "name", ""))
                    em = _safe(getattr(r, "email", ""))
                    vals.append(f"{nm} <{em}>" if em else nm)
            if vals:
                return _join(vals)
    except Exception:
        pass
    # fallback to string field
    return _safe(getattr(msg, field, ""))

def _attachment_names(msg) -> List[str]:
    names = []
    try:
        for att in (msg.attachments or []):
            # extract_msg.Attachment has .longFilename or .shortFilename
            name = _safe(getattr(att, "longFilename", "") or getattr(att, "shortFilename", ""))
            if name:
                names.append(name)
    except Exception:
        pass
    return names

# ---- Core -------------------------------------------------------------------
def convert_to_csv(input_dir: Path, output_csv_path: Path) -> Tuple[int, List[str]]:
    """
    Convert all .msg files in input_dir into a single CSV file at output_csv_path.
    Returns (rows_written, errors).
    """
    input_dir = Path(input_dir)
    msg_files = []
    if input_dir.is_file() and input_dir.suffix.lower() == ".msg":
        msg_files = [input_dir]
    elif input_dir.is_dir():
        msg_files = sorted(input_dir.rglob("*.msg"))
    output_csv_path = Path(output_csv_path)
    output_csv_path.parent.mkdir(parents=True, exist_ok=True)

    errors: List[str] = []
    rows: List[EmailRow] = []

    for msg_path in msg_files:
        try:
            msg = extract_msg.Message(msg_path)
            msg_message_date = _normalize_date(_safe(getattr(msg, "date", "")))
            atts = _attachment_names(msg)

            row = EmailRow(
                file_name=msg_path.name,
                subject=_safe(getattr(msg, "subject", "")),
                from_name=_safe(getattr(msg, "sender", "")),
                from_email=_safe(getattr(msg, "sender_email", "")),
                to=_recipient_list(msg, "to"),
                cc=_recipient_list(msg, "cc"),
                bcc=_recipient_list(msg, "bcc"),
                date_utc=msg_message_date,
                attachments_count=len(atts),
                attachments="; ".join(atts),
                body_text=_safe(getattr(msg, "body", "")),  # plain text body; safe for CSV
            )
            rows.append(row)
        except Exception as e:
            errors.append(f"{msg_path}: {e}")

    # Excel-friendly UTF-8 with BOM
    with open(output_csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[k for k in asdict(rows[0]).keys()] if rows else [f.name for f in EmailRow.__dataclass_fields__.values()],
            quoting=csv.QUOTE_MINIMAL,
        )
        writer.writeheader()
        for r in rows:
            writer.writerow(asdict(r))

    return (len(rows), errors)

# Optional tiny CLI for ad-hoc use:
if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser(description="Batch-convert .msg to a single CSV.")
    p.add_argument("input_dir", type=Path, help="Folder containing .msg files")
    p.add_argument("output_csv", type=Path, help="Path to write CSV (e.g., out/messages.csv)")
    args = p.parse_args()
    count, errs = convert_to_csv(args.input_dir, args.output_csv)
    print(f"Wrote {count} rows -> {args.output_csv}")
    if errs:
        print("Errors:")
        for e in errs:
            print(" -", e)
