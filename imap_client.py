import email as email_lib
import imaplib
from datetime import datetime, timezone
from email.utils import parseaddr, parsedate_to_datetime

from .mail_utils import decode_mime_header, extract_text_body


def is_recent_email(mail_info: dict, init_dt: datetime) -> bool:
    """Check if email date is after init_dt."""
    date_str = mail_info.get("date_raw", "")
    if not date_str:
        return True
    try:
        mail_dt = datetime.fromisoformat(date_str)
        if mail_dt.tzinfo is None:
            mail_dt = mail_dt.replace(tzinfo=timezone.utc)
        if init_dt.tzinfo is None:
            init_dt = init_dt.replace(tzinfo=timezone.utc)
        return mail_dt >= init_dt
    except Exception:
        return True


def _connect(account: dict) -> imaplib.IMAP4 | imaplib.IMAP4_SSL:
    """Create and authenticate an IMAP connection."""
    server = account["imap_server"]
    port = account.get("imap_port", 993)
    use_ssl = account.get("use_ssl", True)

    if use_ssl:
        conn = imaplib.IMAP4_SSL(server, port, timeout=30)
    else:
        conn = imaplib.IMAP4(server, port, timeout=30)

    conn.login(account["email"], account["password"])
    conn.select("INBOX", readonly=True)
    return conn


def _parse_email(msg, uid: bytes, max_body_len: int) -> dict:
    """Parse a single email message into a dict."""
    subject = decode_mime_header(msg.get("Subject", ""))
    from_raw = msg.get("From", "")
    from_name, from_addr = parseaddr(from_raw)
    from_name = decode_mime_header(from_name) if from_name else from_addr

    date_str = msg.get("Date", "")
    try:
        dt = parsedate_to_datetime(date_str)
        date_formatted = dt.strftime("%Y-%m-%d %H:%M")
        date_raw = dt.isoformat()
    except Exception:
        date_formatted = date_str[:25] if date_str else "未知"
        date_raw = ""

    body = extract_text_body(msg, max_body_len)

    return {
        "uid": int(uid),
        "subject": subject or "(无主题)",
        "from_name": from_name,
        "from_addr": from_addr,
        "date": date_formatted,
        "date_raw": date_raw,
        "body": body,
    }


def imap_fetch_new(
    account: dict, last_uid: int, max_body_len: int
) -> tuple[list[dict], int]:
    """Fetch new emails since last_uid. Synchronous — run via asyncio.to_thread."""
    conn = _connect(account)
    new_emails: list[dict] = []
    new_max_uid = last_uid

    try:
        if last_uid == 0:
            status, data = conn.uid("search", "ALL")
            if status == "OK" and data[0]:
                uid_list = data[0].split()
                if uid_list:
                    new_max_uid = int(uid_list[-1])
            return [], new_max_uid

        status, data = conn.uid("search", f"UID {last_uid + 1}:*")
        if status != "OK" or not data[0]:
            return [], last_uid

        uid_list = data[0].split()
        uid_list = [u for u in uid_list if int(u) > last_uid]
        if not uid_list:
            return [], last_uid

        new_max_uid = max(int(u) for u in uid_list)

        for uid in uid_list[-10:]:
            status, msg_data = conn.uid("fetch", uid, "(RFC822)")
            if status != "OK" or not msg_data or not msg_data[0]:
                continue
            raw = msg_data[0][1]
            if not isinstance(raw, bytes):
                continue
            msg = email_lib.message_from_bytes(raw)
            new_emails.append(_parse_email(msg, uid, max_body_len))
    finally:
        try:
            conn.logout()
        except Exception:
            pass

    return new_emails, new_max_uid


def imap_query_since(
    account: dict, since_dt: datetime, max_body_len: int
) -> list[dict]:
    """Query emails since a given date. Synchronous — run via asyncio.to_thread."""
    conn = _connect(account)
    results: list[dict] = []

    try:
        since_str = since_dt.strftime("%d-%b-%Y")
        status, data = conn.uid("search", f"SINCE {since_str}")
        if status != "OK" or not data[0]:
            return []

        uid_list = data[0].split()
        for uid in uid_list[-20:]:
            status, msg_data = conn.uid("fetch", uid, "(RFC822)")
            if status != "OK" or not msg_data or not msg_data[0]:
                continue
            raw = msg_data[0][1]
            if not isinstance(raw, bytes):
                continue
            msg = email_lib.message_from_bytes(raw)
            results.append(_parse_email(msg, uid, max_body_len))
    finally:
        try:
            conn.logout()
        except Exception:
            pass

    return results
