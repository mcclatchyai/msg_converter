# converters/msg_to_eml.py
import extract_msg
from email.message import EmailMessage
from email import message_from_string
from email.utils import format_datetime
from datetime import datetime

def convert_to_eml(msg_path, output_path):
    msg = extract_msg.Message(str(msg_path))
    email = EmailMessage()

    # Basic headers
    email['Subject'] = msg.subject or ""
    email['From'] = msg.sender or ""
    email['To'] = msg.to or ""
    if msg.cc:
        email['Cc'] = msg.cc
    parsed_headers = message_from_string(str(msg.header))
    reply_to = parsed_headers.get("Reply-To")
    if reply_to:
        email['Reply-To'] = reply_to
    if msg.date:
        try:
            parsed_date = datetime.strptime(msg.date, "%a, %d %b %Y %H:%M:%S %z")
            email['Date'] = format_datetime(parsed_date)
        except Exception:
            email['Date'] = msg.date  # fallback

    # Message body
    email.set_content(msg.body or "")

    # Attachments
    for attachment in msg.attachments:
        email.add_attachment(
            attachment.data,
            maintype='application',
            subtype='octet-stream',
            filename=attachment.longFilename or attachment.shortFilename or "attachment"
        )

    # Write to .eml
    with open(output_path, 'wb') as f:
        f.write(email.as_bytes())