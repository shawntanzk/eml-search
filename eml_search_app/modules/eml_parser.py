"""Parse .eml files into structured dicts."""
import email
import hashlib
import re
from email.header import decode_header, make_header
from email.utils import parseaddr, parsedate_to_datetime
from html.parser import HTMLParser
from pathlib import Path
from typing import Optional


class _HTMLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self.reset()
        self._fed: list[str] = []

    def handle_data(self, d: str) -> None:
        self._fed.append(d)

    def get_text(self) -> str:
        return " ".join(self._fed)


def _strip_html(html_text: str) -> str:
    stripper = _HTMLStripper()
    stripper.feed(html_text)
    return stripper.get_text()


def _decode_header_str(raw: str) -> str:
    if not raw:
        return ""
    try:
        return str(make_header(decode_header(raw)))
    except Exception:
        return raw


def _parse_address_list(header_val: str) -> list[dict]:
    if not header_val:
        return []
    result = []
    for part in header_val.split(","):
        part = part.strip()
        name, addr = parseaddr(_decode_header_str(part))
        addr = addr.lower()
        if addr:
            result.append({"name": name or addr.split("@")[0], "email": addr})
    return result


def _decode_payload(part) -> str:
    payload = part.get_payload(decode=True)
    if not payload:
        return ""
    charset = part.get_content_charset() or "utf-8"
    return payload.decode(charset, errors="replace")


def parse_eml(file_path: str) -> Optional[dict]:
    """Parse an .eml file and return a structured dict, or None on failure."""
    try:
        path = Path(file_path)
        file_id = hashlib.md5(str(path.resolve()).encode()).hexdigest()

        with open(file_path, "rb") as f:
            msg = email.message_from_bytes(f.read())

        subject = _decode_header_str(msg.get("Subject", ""))

        from_raw = _decode_header_str(msg.get("From", ""))
        sender_name, sender_email = parseaddr(from_raw)
        sender_email = sender_email.lower()
        sender_name = sender_name or (sender_email.split("@")[0] if sender_email else "")

        recipients = _parse_address_list(msg.get("To", ""))
        cc = _parse_address_list(msg.get("CC", ""))

        date_str = msg.get("Date", "")
        try:
            date = parsedate_to_datetime(date_str).isoformat() if date_str else None
        except Exception:
            date = None

        message_id = msg.get("Message-ID", "").strip("<>").strip()
        in_reply_to = msg.get("In-Reply-To", "").strip("<>").strip()

        refs_raw = msg.get("References", "").strip()
        refs = [r.strip("<>").strip() for r in refs_raw.split() if r.strip()]
        thread_id = refs[0] if refs else (in_reply_to or message_id)

        body_text = ""
        body_html = ""
        attachments: list[str] = []

        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                disposition = str(part.get("Content-Disposition", ""))

                if "attachment" in disposition:
                    fname = part.get_filename()
                    if fname:
                        attachments.append(_decode_header_str(fname))
                elif content_type == "text/plain" and not body_text:
                    body_text = _decode_payload(part)
                elif content_type == "text/html" and not body_html:
                    body_html = _decode_payload(part)
        else:
            content_type = msg.get_content_type()
            body_text = _decode_payload(msg)
            if content_type == "text/html":
                body_html = body_text
                body_text = ""

        if not body_text and body_html:
            body_text = _strip_html(body_html)

        body_text = re.sub(r"\s+", " ", body_text).strip()

        return {
            "id": file_id,
            "file_path": str(path.resolve()),
            "message_id": message_id,
            "subject": subject,
            "sender_name": sender_name,
            "sender_email": sender_email,
            "recipients": recipients,
            "cc": cc,
            "date": date,
            "body_text": body_text,
            "has_attachments": len(attachments) > 0,
            "attachment_names": attachments,
            "thread_id": thread_id,
            "in_reply_to": in_reply_to,
        }
    except Exception:
        return None
