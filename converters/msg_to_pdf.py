# converters/msg_to_pdf.py
import extract_msg
from weasyprint import HTML
from pathlib import Path
from email import message_from_string
import html
import os
import base64
import re

def remove_page_rules(html: str) -> str:
    # Remove all @page blocks (even multiline)
    html = re.sub(r'@page\s+[^{]+{[^}]*}', '', html, flags=re.IGNORECASE | re.DOTALL)
    # Remove all 'page: something;' declarations
    html = re.sub(r'page\s*:\s*[^;{]+;', '', html, flags=re.IGNORECASE)
    return html

def convert_to_pdf(msg_path, output_path):
    msg = extract_msg.Message(str(msg_path))

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

    if msg.attachments:
        attachments_dir = Path(output_path).with_suffix('.attachments')
        attachments_dir.mkdir(exist_ok=True)
        for attachment in msg.attachments:
            name = attachment.longFilename or attachment.shortFilename or "attachment"
            path = attachments_dir / name
            with open(path, "wb") as f:
                f.write(attachment.data)

            ext = Path(name).suffix.lower().replace('.', '')
            rel_path = path.relative_to(Path(output_path).parent)

            if ext in {"png", "jpg", "jpeg", "gif"}:
                mime = f"image/{'jpeg' if ext == 'jpg' else ext}"
                b64data = base64.b64encode(attachment.data).decode('utf-8')
                inline_images.add(name)
                cid_url = f"cid:{name}"
                data_url = f"data:{mime};base64,{b64data}"
                if msg.htmlBody:
                    decoded_html = msg.htmlBody.decode('utf-8', errors='ignore')
                    decoded_html = decoded_html.replace(cid_url, data_url)
                    msg.htmlBody = decoded_html.encode('utf-8')
            else:
                size_kb = round(os.path.getsize(path) / 1024, 1)
                link = f"<a href='{rel_path}'>{html.escape(name)}</a> ({size_kb} KB)"
                external_files.append(link)

        if external_files:
            attachments_str = ', '.join(external_files)

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