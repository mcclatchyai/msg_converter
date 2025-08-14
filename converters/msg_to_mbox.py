# converters/msg_to_mbox.py
import mailbox
from pathlib import Path

def convert_to_mbox(eml_paths, mbox_path):
    mbox = mailbox.mbox(str(mbox_path))
    for eml_path in eml_paths:
        with open(eml_path, 'rb') as f:
            msg = mailbox.mboxMessage(f)
            mbox.add(msg)
    mbox.flush()
