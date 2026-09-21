import hashlib
import imaplib
import json
import logging
import smtplib
import ssl
import time
from contextlib import contextmanager, suppress
from dataclasses import dataclass
from datetime import UTC, datetime
from email import policy
from email.message import EmailMessage
from email.parser import BytesParser
from email.utils import format_datetime, make_msgid, parseaddr
from pathlib import Path
from typing import Iterable
from uuid import uuid4

from .config import Settings
from .models import ForecastRequest, ForecastResponse

logger = logging.getLogger(__name__)


def _send_now(settings: Settings, message: EmailMessage) -> None:
    """Submit directly to SMTP, independent of The Bat! and its outbox."""
    if not message.get("Message-ID"):
        message["Message-ID"] = make_msgid()
    client = _open_smtp(settings)
    try:
        refused = client.send_message(message)
        if refused:
            raise smtplib.SMTPRecipientsRefused(refused)
        logger.info("SMTP accepted message %s", message["Message-ID"])
    finally:
        # A QUIT failure after acceptance must not prompt a duplicate send.
        try:
            client.quit()
        except (OSError, smtplib.SMTPException):
            client.close()


@dataclass(frozen=True)
class ImapResponseEnvelope:
    uid: bytes
    response: ForecastResponse


@dataclass(frozen=True)
class ImapInputEnvelope:
    uid: bytes
    sender: str
    subject: str
    filename: str
    content: bytes
    uidvalidity: str = ""


def check_connections(settings: Settings) -> dict[str, str]:
    """Authenticate to IMAP and SMTP without sending or changing any mail."""
    imap = _open_imap(settings)
    try:
        imap.noop()
    finally:
        imap.logout()

    smtp = _open_smtp(settings)
    try:
        smtp.noop()
    finally:
        smtp.quit()
    return {"imap": "ok", "smtp": "ok"}


def run_self_test(settings: Settings, timeout_seconds: int = 30) -> dict[str, str]:
    """Send one JSON message to the same mailbox and read it back over IMAP."""
    test_id = str(uuid4())
    subject = f"[MAIL_SELF_TEST] {test_id}"
    expected = {"test_id": test_id, "status": "mail-roundtrip-test"}

    message = EmailMessage()
    message["From"] = settings.MAILBOX_NAME
    message["To"] = settings.MAILBOX_NAME
    message["Subject"] = subject
    message["Date"] = format_datetime(datetime.now(UTC))
    message.set_content("Автоматическая проверка SMTP -> IMAP.")
    message.add_attachment(
        json.dumps(expected, ensure_ascii=False).encode("utf-8"),
        maintype="application",
        subtype="json",
        filename="mail_self_test.json",
    )

    _send_now(settings, message)

    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        imap = _open_imap(settings)
        try:
            imap.select("INBOX")
            status, data = imap.uid("search", None, "SUBJECT", f'"{subject}"')
            if status == "OK" and data[0]:
                uid = data[0].split()[-1]
                status, raw = imap.uid("fetch", uid, "(RFC822)")
                if status != "OK" or not raw or not isinstance(raw[0], tuple):
                    raise RuntimeError(
                        "Test message was found but could not be fetched"
                    )
                parsed = BytesParser(policy=policy.default).parsebytes(raw[0][1])
                attachments = list(parsed.iter_attachments())
                if len(attachments) != 1:
                    raise RuntimeError(
                        "Expected exactly one attachment in test message"
                    )
                actual = json.loads(
                    attachments[0].get_payload(decode=True).decode("utf-8-sig")
                )
                if actual != expected:
                    raise RuntimeError(
                        "JSON attachment does not match the sent payload"
                    )
                imap.uid("store", uid, "+FLAGS", "(\\Seen)")
                return {
                    "smtp": "sent",
                    "imap": "received",
                    "json": "verified",
                    "test_id": test_id,
                }
        finally:
            imap.logout()
        time.sleep(2)
    raise TimeoutError(
        f"Test message was not received within {timeout_seconds} seconds"
    )


def _open_imap(settings: Settings):
    if settings.IMAP_USE_SSL:
        client = imaplib.IMAP4_SSL(
            settings.IMAP_SERVER,
            settings.IMAP_PORT,
            ssl_context=ssl.create_default_context(),
            timeout=settings.MAIL_TIMEOUT_SECONDS,
        )
    else:
        client = imaplib.IMAP4(
            settings.IMAP_SERVER,
            settings.IMAP_PORT,
            timeout=settings.MAIL_TIMEOUT_SECONDS,
        )
    try:
        if not settings.IMAP_USE_SSL:
            client.starttls(ssl_context=ssl.create_default_context())
        client.login(settings.MAILBOX_NAME, settings.MAILBOX_PASSWORD)
    except Exception:
        with suppress(Exception):
            client.shutdown()
        raise
    return client


def _open_smtp(settings: Settings):
    if not settings.SMTP_USE_SSL and not settings.SMTP_USE_STARTTLS:
        raise ValueError("SMTP requires TLS before authentication")
    context = ssl.create_default_context()
    if settings.SMTP_USE_SSL:
        client = smtplib.SMTP_SSL(
            settings.SMTP_SERVER,
            settings.SMTP_PORT,
            context=context,
            timeout=settings.MAIL_TIMEOUT_SECONDS,
        )
    else:
        client = smtplib.SMTP(
            settings.SMTP_SERVER,
            settings.SMTP_PORT,
            timeout=settings.MAIL_TIMEOUT_SECONDS,
        )
    try:
        client.ehlo()
        if not settings.SMTP_USE_SSL:
            client.starttls(context=context)
            client.ehlo()
        client.login(settings.MAILBOX_NAME, settings.MAILBOX_PASSWORD)
    except Exception:
        client.close()
        raise
    return client


@contextmanager
def _inbox(settings: Settings, *, readonly: bool = True):
    client = _open_imap(settings)
    try:
        status, _ = client.select("INBOX", readonly=readonly)
        if status != "OK":
            raise RuntimeError("Cannot select INBOX")
        yield client
    finally:
        with suppress(OSError, imaplib.IMAP4.error):
            client.logout()


def _iter_messages(settings: Settings, prefix: str):
    """Read without changing Seen; acknowledge separately after processing."""
    if not prefix or any(c in prefix for c in '\r\n"\\') or not prefix.isascii():
        raise ValueError(
            "Subject prefix must be nonempty ASCII without quotes or backslashes"
        )
    with _inbox(settings) as client:
        validity = client.response("UIDVALIDITY")[1]
        if not validity or not validity[0]:
            raise RuntimeError("Server did not return UIDVALIDITY")
        uidvalidity = validity[0].decode("ascii")
        status, data = client.uid("search", None, "UNSEEN", "SUBJECT", f'"{prefix}"')
        if status != "OK":
            raise RuntimeError("IMAP search failed")
        for uid in (data[0] or b"").split():
            status, raw = client.uid("fetch", uid, "(BODY.PEEK[])")
            if status != "OK":
                raise RuntimeError("IMAP fetch failed")
            payload = next((item[1] for item in raw if isinstance(item, tuple)), None)
            if payload is None:
                raise RuntimeError("IMAP response contains no message body")
            message = BytesParser(policy=policy.default).parsebytes(payload)
            if str(message.get("Subject", "")).startswith(prefix):
                yield uid, uidvalidity, message


def send_forecast_request(settings: Settings, request: ForecastRequest) -> None:
    recipient = settings.INTERNAL_OUTLOOK_EMAIL.strip()
    if not recipient:
        raise ValueError("INTERNAL_OUTLOOK_EMAIL is required")
    message = EmailMessage()
    message["From"] = settings.MAILBOX_NAME
    message["To"] = recipient
    message["Reply-To"] = settings.MAILBOX_NAME
    message["Subject"] = f"{settings.N8N_REQUEST_SUBJECT_PREFIX} {request.request_id}"
    message["Date"] = format_datetime(datetime.now(UTC))
    message.set_content("JSON-запрос прогноза находится во вложении.")
    message.add_attachment(
        request.model_copy(update={"reply_to": settings.MAILBOX_NAME})
        .model_dump_json(indent=2)
        .encode("utf-8"),
        maintype="application",
        subtype="json",
        filename="forecast_request.json",
    )
    _send_now(settings, message)


def iter_input_workbooks(settings: Settings) -> Iterable[ImapInputEnvelope]:
    """Yield input attachments without acknowledging their source message."""
    for uid, validity, message in _iter_messages(
        settings, settings.INPUT_SUBJECT_PREFIX
    ):
        sender = parseaddr(str(message.get("From", "")))[1].strip().lower()
        if not sender or (
            settings.TRUSTED_REQUESTERS and sender not in settings.TRUSTED_REQUESTERS
        ):
            logger.warning("Ignored input sender uid=%s", uid.decode())
            continue
        for part in message.iter_attachments():
            filename = (part.get_filename() or "").replace("\\", "/").split("/")[-1]
            if Path(filename).suffix.lower() != ".xlsx":
                continue
            content = part.get_payload(decode=True)
            if not content:
                raise ValueError("Empty XLSX attachment")
            # Never use an untrusted attachment filename as a filesystem path.
            yield ImapInputEnvelope(
                uid,
                sender,
                str(message.get("Subject", "")),
                "input-" + hashlib.sha256(filename.encode()).hexdigest()[:16] + ".xlsx",
                content,
                validity,
            )


def iter_forecast_responses(settings: Settings) -> Iterable[ImapResponseEnvelope]:
    expected_sender = (
        (settings.N8N_RESPONSE_SENDER or settings.INTERNAL_OUTLOOK_EMAIL)
        .strip()
        .lower()
    )
    if not expected_sender:
        raise ValueError("N8N_RESPONSE_SENDER or INTERNAL_OUTLOOK_EMAIL is required")
    for uid, _, message in _iter_messages(
        settings, settings.N8N_RESPONSE_SUBJECT_PREFIX
    ):
        sender = parseaddr(str(message.get("From", "")))[1].strip().lower()
        if sender != expected_sender:
            logger.warning("Ignored unexpected response sender uid=%s", uid.decode())
            continue
        for part in message.iter_attachments():
            if Path(part.get_filename() or "").suffix.lower() != ".json":
                continue
            try:
                content = part.get_payload(decode=True)
                response = ForecastResponse.model_validate_json(
                    content.decode("utf-8-sig")
                )
            except (ValueError, AttributeError):
                logger.warning("Ignored invalid JSON response uid=%s", uid.decode())
                continue
            yield ImapResponseEnvelope(uid, response)


def mark_processed(settings: Settings, uid: bytes) -> None:
    with _inbox(settings, readonly=False) as client:
        status, _ = client.uid("store", uid, "+FLAGS", "(\\Seen)")
        if status != "OK":
            raise RuntimeError("Could not mark message processed")


def send_result_workbook(
    settings: Settings,
    recipient: str,
    request_id: str,
    workbook_path: str | Path,
) -> None:
    message = EmailMessage()
    message["From"] = settings.MAILBOX_NAME
    message["To"] = recipient
    message["Subject"] = f"[FORECAST_RESULT] {request_id}"
    message["Date"] = format_datetime(datetime.now(UTC))
    message.set_content("Результат прогноза находится во вложении.")
    path = Path(workbook_path)
    message.add_attachment(
        path.read_bytes(),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=path.name,
    )
    _send_now(settings, message)
