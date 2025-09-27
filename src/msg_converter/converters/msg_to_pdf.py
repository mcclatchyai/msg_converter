# converters/msg_to_pdf.py
import extract_msg
from weasyprint import HTML
from pathlib import Path
from email import message_from_string
import html
import os
import base64
import re

_filename_safe = re.compile(r"[^\w\-.() ]+")

def sanitize_filename(name: str) -> str:
    # drop directory parts and scrub odd chars
    base = Path(name).name
    base = _filename_safe.sub("_", base).strip()
    return base or "attachment"

def remove_page_rules(html: str) -> str:
  # Remove all @page blocks (even multiline)
  html = re.sub(r'@page\s+[^{]+{[^}]*}', '', html, flags=re.IGNORECASE | re.DOTALL)
  # Remove all 'page: something;' declarations
  html = re.sub(r'page\s*:\s*[^;{]+;', '', html, flags=re.IGNORECASE)
  return html

def write_bytes(path: Path, data: bytes) -> Path:
    path = unique_path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "wb") as f:
        f.write(data)
    return path

def unique_path(p: Path) -> Path:
    if not p.exists():
        return p
    i = 1
    while True:
        cand = p.with_name(f"{p.stem} ({i}){p.suffix}")
        if not cand.exists():
            return cand
        i += 1

def is_msg_attachment(att) -> bool:
    fn = (getattr(att, "longFilename", None) or getattr(att, "shortFilename", None) or "")
    if fn.lower().endswith(".msg"):
        return True
    data = getattr(att, "data", None)
    # extract_msg returns a MSGFile-like object for embedded emails
    if data is not None and not isinstance(data, (bytes, bytearray)) \
       and data.__class__.__name__.lower().startswith("msg"):
        return True
    return False

def core_ext_from_double_ext(name: str) -> str | None:
    # e.g., "x.doc.msg" -> ".doc"
    parts = name.lower().split(".")
    if len(parts) >= 2 and parts[-1] == "msg":
        ext = parts[-2]
        if ext in {"doc", "xls", "ppt", "rtf", "pdf", "docx", "xlsx", "pptx"}:
            return "." + ext
    return None

def sniff_archive_office_ext(path: Path, fallback: str | None = None) -> str | None:
    try:
        with open(path, "rb") as f:
            head = f.read(8)
        if head.startswith(b"PK\x03\x04"):
            # OOXML container; prefer the core ext if present
            return fallback if fallback in {".docx", ".xlsx", ".pptx"} else fallback or ".zip"
        if head.startswith(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"):
            # OLE compound (legacy Office: doc/xls/ppt or MSG)
            return fallback if fallback in {".doc", ".xls", ".ppt"} else fallback
    except Exception:
        pass
    return fallback

def sniff_ext(path: Path) -> str | None:
    try:
        with open(path, "rb") as f:
            head = f.read(8)
        if head.startswith(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"):
            return ".msg"  # OLE compound; good heuristic for embedded emails here
        if head.startswith(b"PK\x03\x04"):
            return ".zip"  # docx/xlsx/pptx containers
    except Exception:
        pass
    return None
  
def save_via_attachment(att, target_dir: Path, target_name: str) -> Path | None:
    """
    Robustly save any attachment (including embedded .msg) into target_dir.
    - Forces embedded .msg to be written as a .msg file (not MSGFile.save()).
    - Works across extract_msg versions/param signatures.
    - Falls back to bytes if needed.
    """
    target_dir.mkdir(parents=True, exist_ok=True)
    target_name = sanitize_filename(target_name)
    candidate = unique_path(target_dir / target_name)

    # Snapshot directory to detect what actually got created
    before = set(os.listdir(target_dir))
    # Try preferred signatures first
    try:
        # Directory + explicit filename, and force embedded MSG as bytes-on-disk
        att.save(customPath=str(target_dir),
                 customFilename=candidate.name,
                 extractEmbedded=True)
    except TypeError:
        # Older sigs may not accept customFilename; let library choose a name
        att.save(customPath=str(target_dir), extractEmbedded=True)
    except Exception:
        # ---- Byte fallbacks ----
        data = getattr(att, "data", None)
        if isinstance(data, (bytes, bytearray)):
            return write_bytes(candidate, data)
        for attr in ("as_bytes", "toBytes", "data"):
            if hasattr(data, attr):
                val = getattr(data, attr)
                buf = val() if callable(val) else val
                if isinstance(buf, (bytes, bytearray)):
                    return write_bytes(candidate, buf)
        return None  # nothing worked

    # If the library wrote exactly our candidate, great.
    if candidate.exists():
        return candidate

    # Otherwise, find what new file(s) appeared and return the best match.
    after = set(os.listdir(target_dir))
    created = list(after - before)
    if not created:
        # Some versions return an existing name; try the attachment's own names
        for guess_name in filter(None, [att.longFilename, att.shortFilename]):
            guess = target_dir / sanitize_filename(guess_name)
            if guess.exists():
                return guess
        return None

    # Prefer a file with the same extension as our candidate, else first new file.
    cand_ext = candidate.suffix.lower()
    for nm in created:
        if Path(nm).suffix.lower() == cand_ext:
            return target_dir / nm
    return target_dir / created[0]


def convert_to_pdf(msg_path, output_path):
    msg = extract_msg.Message(str(msg_path))

    #print(msg_path)

    parsed_headers = message_from_string(str(msg.header))
    from_ = parsed_headers.get("From", msg.sender or "(Unknown Sender)")
    to = parsed_headers.get("To", msg.to or "(No Recipient)")
    cc = parsed_headers.get("Cc", msg.cc or "")
    reply_to = parsed_headers.get("Reply-To", "")
    date = parsed_headers.get("Date", msg.date or "")
    subject = parsed_headers.get("Subject", msg.subject or "(No Subject)")

    attachments_str = ""
    inline_images = set()
    external_files = []

    html_raw = msg.htmlBody.decode("utf-8", errors="ignore") if msg.htmlBody else ""
    # capture cid:XYZ tokens from html (e.g., src="cid:image001.png@..."):
    cid_refs = set(re.findall(r'cid:([^"\'\s>]+)', html_raw, flags=re.IGNORECASE))

    if msg.attachments:
        attachments_dir = Path(output_path).with_suffix(".attachments")
        attachments_dir.mkdir(exist_ok=True)

        def rel_to_doc(p: Path) -> str | None:
            try:
                return str(p.relative_to(Path(output_path).parent))
            except Exception:
                return None

        for attachment in msg.attachments:
            raw_name = attachment.longFilename or attachment.shortFilename or "attachment"
            name = sanitize_filename(raw_name)
            ext = Path(name).suffix.lower()

            # --- Treat as embedded .msg even if the filename lacks .msg ---
            embedded_email = is_msg_attachment(attachment)
            if embedded_email and ext != ".msg":
                name = name + ".msg"
                ext = ".msg"

            # ---------- Case A: embedded email (.msg) -> save then convert ----------
            if embedded_email:
                saved_msg = save_via_attachment(attachment, attachments_dir, name)
                if not (saved_msg and saved_msg.exists()):
                    external_files.append(f"{html.escape(name)} (not saved)")
                    continue

                # Try to parse as real MSG; if that fails, treat as a normal file
                is_real_msg = True
                try:
                    _probe = extract_msg.Message(str(saved_msg))
                except Exception:
                    is_real_msg = False

                if not is_real_msg:
                    # Not actually a .msg → rename to core ext (e.g., .doc) or sniffed ext
                    core = core_ext_from_double_ext(raw_name)
                    guessed = sniff_archive_office_ext(saved_msg, fallback=core)
                    if guessed and saved_msg.suffix.lower() != guessed:
                        try:
                            newp = unique_path(saved_msg.with_suffix(guessed))
                            saved_msg.rename(newp)
                            saved_msg = newp
                        except Exception:
                            pass

                    href = rel_to_doc(saved_msg)
                    size_kb = round(saved_msg.stat().st_size / 1024, 1)
                    label = html.escape(saved_msg.name)
                    external_files.append(
                        f"<a href='{href}'>{label}</a> ({size_kb} KB)" if href else f"{label} ({size_kb} KB)"
                    )
                    continue  # done with this attachment

                # --- Real MSG → convert to PDF ---
                inner_subject = None
                try:
                    inner = extract_msg.Message(str(saved_msg))
                    inner_subject = inner.subject
                except Exception:
                    pass

                pretty_stem = sanitize_filename(inner_subject) if inner_subject else saved_msg.stem
                pdf_path = unique_path(attachments_dir / f"{pretty_stem}.pdf")
                try:
                    convert_to_pdf(str(saved_msg), str(pdf_path))
                    href = rel_to_doc(pdf_path)
                    size_kb = round(pdf_path.stat().st_size / 1024, 1) if pdf_path.exists() else "?"
                    # Display as .msg but link to PDF
                    display_label = html.escape(f"{pretty_stem}.msg")
                    external_files.append(
                        f"<a href='{href}'>{display_label}</a> ({size_kb} KB)" if href else f"{display_label} ({size_kb} KB)"
                    )
                except Exception as e:
                    display_label = html.escape(f"{pretty_stem}.msg")
                    external_files.append(f"{display_label} (conversion failed: {html.escape(str(e))})")
                continue


            # ---------- Case B: regular attachments (images/docs) ----------
            # Detect inline by CID match
            cid_candidates = set()
            for attr in ("contentId", "cid", "content_id"):
                val = getattr(attachment, attr, None)
                if val:
                    cid_candidates.add(str(val).strip("<>"))
            cid_candidates.add(name)  # filenames sometimes used directly in cid:NAME

            inline_hit = any(c in cid_refs for c in cid_candidates)

            if inline_hit:
                # Embed inline if we have bytes; do NOT save or list as attachment
                data = getattr(attachment, "data", None)
                if isinstance(data, (bytes, bytearray)):
                    ext_no_dot = Path(name).suffix.lower().lstrip(".")
                    if ext_no_dot in {"png", "jpg", "jpeg", "gif"}:
                        mime = f"image/{'jpeg' if ext_no_dot == 'jpg' else ext_no_dot}"
                        b64data = base64.b64encode(data).decode("utf-8")
                        cid_urls = {f"cid:{c}" for c in cid_candidates}  # replace any variant
                        if msg.htmlBody:
                            decoded_html = msg.htmlBody.decode("utf-8", errors="ignore")
                            for cid_url in cid_urls:
                                decoded_html = decoded_html.replace(cid_url, f"data:{mime};base64,{b64data}")
                            msg.htmlBody = decoded_html.encode("utf-8")
                # Skip saving/listing entirely
                continue

            # Not inline → treat as a real attachment (save + link)
            saved_file = save_via_attachment(attachment, attachments_dir, name)
            if saved_file and saved_file.exists():
                href = rel_to_doc(saved_file)
                size_kb = round(saved_file.stat().st_size / 1024, 1)
                label = html.escape(saved_file.name)
                external_files.append(
                    f"<a href='{href}'>{label}</a> ({size_kb} KB)" if href else f"{label} ({size_kb} KB)"
                )
            else:
                external_files.append(f"{html.escape(name)} (not saved)")

        if external_files:
            attachments_str = ", ".join(external_files)

    if msg.htmlBody:
      raw_body = msg.htmlBody.decode('utf-8', errors='ignore')
      # Strip <html>, <head>, <body> tags only
      body = re.sub(r'</?(html|head|body)[^>]*>', '', raw_body, flags=re.IGNORECASE).strip()
    elif msg.rtfBody:
      body = html.escape(msg.rtfBody)
    else:
      body = html.escape(msg.body or '')

    cleaned_body = remove_page_rules(body)

    html_content = f"""
      <html>
      <head>
        <meta charset='utf-8'>
        <style>
          * {{ page: auto !important; }}
          @page {{ size: 8.5in 11in; margin: 0.5in; }}
          body {{ font-family: sans-serif; margin: 0; box-sizing: border-box; }}
          .meta-table {{ 
            font-size: 0.9em;
            border-collapse: collapse;
            margin-bottom: 0.75em;
            width: 100%;
            table-layout: auto;
            break-after: avoid;
          }}
          .meta-table td {{
            vertical-align: top;
            padding: 2px 6px 2px 0;
          }}
          .meta-label {{
            font-weight: bold;
            white-space: nowrap;
            width: 1%;
          }}
          hr {{
            border: none;
            border-top: 1px solid #ccc;
            margin: 6px 0 12px 0;
            break-inside: avoid; 
            break-after: avoid; 
          }}
          .email-body-container {{
            display: block;
            position: relative;
            /*overflow: hidden;*/
            width: 100%;
            page-break-inside: auto;
          }}
          .email-body-inner {{
            position: relative;
            width: 100%;
          }}
          .email-normalized * {{
            max-width: 100% !important;
            box-sizing: border-box !important;
          }}

          .email-normalized table {{
            width: 100% !important;
            /*table-layout: fixed !important;*/
          }}

          .email-normalized {{
            width: 100%;
            margin: 0;
            padding: 0;
          }}

          a {{
            text-decoration: none;
            color: #0645ad;
          }}
        </style>
      </head>
      <body>
        <div class="content-block">
          <table class="meta-table">
            <tr><td class="meta-label">From:</td><td>{html.escape(from_)}</td></tr>
            <tr><td class="meta-label">To:</td><td>{html.escape(to)}</td></tr>
            {f"<tr><td class='meta-label'>Cc:</td><td>{html.escape(cc)}</td></tr>" if cc else ''}
            {f"<tr><td class='meta-label'>Reply-To:</td><td>{html.escape(reply_to)}</td></tr>" if reply_to else ''}
            {f"<tr><td class='meta-label'>Date:</td><td>{html.escape(date)}</td></tr>" if date else ''}
            <tr><td class="meta-label">Subject:</td><td>{html.escape(subject)}</td></tr>
            {f"<tr><td class='meta-label'>Attachments:</td><td>{attachments_str}</td></tr>" if attachments_str else ''}
          </table>
          <hr>
          <div class="email-body-container">
            <div class="email-body-inner email-normalized">
              {cleaned_body}
            </div>
          </div>
        </div>
      </body>
      </html>
      """

    # DEBUG
    # Save the HTML to inspect rendering issues
    # html_path = Path(output_path).with_suffix(".html")
    # with open(html_path, "w", encoding="utf-8") as f:
    #     f.write(html_content)

    HTML(string=html_content, base_url=str(Path(output_path).parent)).write_pdf(str(output_path))

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 3:
        print("Usage: python -m converters.msg_to_pdf <input.msg> <output.pdf>")
        sys.exit(1)

    msg_path = sys.argv[1]
    output_path = sys.argv[2]

    convert_to_pdf(msg_path, output_path)