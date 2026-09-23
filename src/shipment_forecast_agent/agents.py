"""Independent request and result processors sharing a CSV registry."""

import hashlib
import json
import logging
import re
import time
from datetime import UTC, datetime
from email.message import EmailMessage
from email.utils import format_datetime, parseaddr

from .internet_mail import (
    _iter_messages,
    _send_now,
    iter_input_workbooks,
    mark_processed,
    send_result_workbook,
)
from .logging_setup import configure_agent_logging
from .registry import Registry, locked
from .state import submit_once, write_json
from .template_excel import fill_forecasts, request_data

logger = logging.getLogger(__name__)
UID_PATTERN = r"RQ-[a-f0-9]{32}"


def process_requests(settings):
    configure_agent_logging(settings, "request")
    logger.info("Starting request polling cycle")
    registry = Registry(settings.WORK_DIR)
    incoming = settings.input_mailbox()
    processed, failed, acknowledged = [], set(), set()
    with locked(settings.WORK_DIR / "request-agent.lock"):
        selected_uid = None
        for item in iter_input_workbooks(incoming):
            # One invocation consumes exactly one source email. All XLSX
            # attachments of that email belong to the same atomic input unit.
            if selected_uid is None:
                selected_uid = item.uid
            elif item.uid != selected_uid:
                break
            try:
                identity = "\0".join(
                    [
                        incoming.MAILBOX_NAME.lower(),
                        item.uidvalidity,
                        item.uid.decode(),
                        item.filename,
                        hashlib.sha256(item.content).hexdigest(),
                    ]
                )
                key = "RQ-" + hashlib.sha256(identity.encode()).hexdigest()[:32]
                directory = settings.WORK_DIR / "jobs" / key
                directory.mkdir(parents=True, exist_ok=True)
                source = directory / "source.xlsx"
                record = registry.get(key)
                if record and record["status"] in ("prepared", "sent", "completed"):
                    acknowledged.add(item.uid)
                    continue
                if not source.exists():
                    source.write_bytes(item.content)
                payload_path = directory / "request.json"
                payload = request_data(source, settings, key)
                write_json(payload_path, payload)
                registry.put(
                    {
                        "request_id": key,
                        "mailbox": incoming.MAILBOX_NAME,
                        "uidvalidity": item.uidvalidity,
                        "mail_uid": item.uid.decode(),
                        "filename": item.original_filename or item.filename,
                        "source_path": str(source.resolve()),
                        "sender": item.sender,
                        "status": "prepared",
                        "request_received_at": (record or {}).get("request_received_at")
                        or item.received_at,
                    }
                )
                logger.info("Request %s prepared and queued", key)
                acknowledged.add(item.uid)
                processed.append(key)
            except Exception:
                failed.add(item.uid)
                logger.exception(
                    "Request processing failed for mail UID %s", item.uid.decode()
                )
        for uid in acknowledged - failed:
            mark_processed(incoming, uid)
    return processed


def _send_prepared_request(settings, registry, record):
    """Send one prepared JSON and atomically advance its queue state."""
    if not settings.INTERNAL_OUTLOOK_EMAIL:
        raise ValueError("INTERNAL_OUTLOOK_EMAIL is required")
    key = record["request_id"]
    if not re.fullmatch(UID_PATTERN, key):
        raise ValueError("Invalid queued request ID")
    directory = settings.WORK_DIR / "jobs" / key
    payload_path = directory / "request.json"
    payload = json.loads(payload_path.read_text(encoding="utf-8"))
    if payload.get("mail_id") != key:
        raise ValueError("Queued JSON mail_id does not match CSV request_id")

    def send():
        message = EmailMessage()
        message["From"] = settings.MAILBOX_NAME
        message["To"] = settings.INTERNAL_OUTLOOK_EMAIL
        message["Reply-To"] = settings.MAILBOX_NAME
        message["Subject"] = f"{settings.N8N_REQUEST_SUBJECT_PREFIX} {key}"
        message["Date"] = format_datetime(datetime.now(UTC))
        message.set_content(f"Запрос прогноза. Верните mail_id {key} в JSON ответа.")
        message.add_attachment(
            json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            maintype="application",
            subtype="json",
            filename="forecast_request.json",
        )
        return _send_now(settings, message)

    submission = directory / "request-submission.json"
    submit_once(submission, send)
    receipt = json.loads(submission.read_text(encoding="utf-8"))
    registry.put(
        {
            "request_id": key,
            "status": "sent",
            "outlook_sent_at": receipt.get("accepted_at", ""),
        }
    )
    logger.info("Queued request %s accepted by SMTP", key)
    return key


def correlate(subject, payload):
    subject_keys = set(
        re.findall(r"(?<![A-Za-z0-9-])" + UID_PATTERN + r"(?![A-Za-z0-9-])", subject)
    )
    keys = set()
    objects = payload if isinstance(payload, list) else [payload]
    if isinstance(payload, dict):
        for field in ("data", "forecasts"):
            if isinstance(payload.get(field), list):
                objects += [r for r in payload[field] if isinstance(r, dict)]
    for obj in objects:
        if isinstance(obj, dict):
            for field in ("mail_id",):
                if field in obj:
                    value = str(obj[field])
                    if not re.fullmatch(UID_PATTERN, value):
                        raise ValueError("Invalid response mail_id")
                    keys.add(value)
    if len(keys) != 1:
        raise ValueError("Response must contain exactly one consistent mail_id in JSON")
    if subject_keys and subject_keys != keys:
        raise ValueError("Subject ID conflicts with JSON mail_id")
    return keys.pop()


def forecast_rows(payload):
    if isinstance(payload, list):
        result = payload
    elif isinstance(payload, dict):
        if payload.get("status", "ok") != "ok":
            raise ValueError("n8n returned an error status")
        result = payload.get("data", payload.get("forecasts", [payload]))
        if "row_count" in payload and payload["row_count"] != len(result):
            raise ValueError("row_count does not match data")
    else:
        raise TypeError("Invalid JSON shape")
    if (
        not isinstance(result, list)
        or not result
        or not all(isinstance(r, dict) for r in result)
    ):
        raise ValueError("Response contains no forecast rows")
    return result


def _handled_response_messages(path):
    if not path.exists():
        return set()
    payload = json.loads(path.read_text(encoding="utf-8"))
    values = payload.get("handled", [])
    if not isinstance(values, list) or not all(isinstance(v, str) for v in values):
        raise ValueError("Invalid result mail ledger")
    return set(values)


def _remember_response(path, handled, token):
    handled.add(token)
    write_json(path, {"handled": sorted(handled)})


def process_results(settings):
    configure_agent_logging(settings, "result")
    logger.info("Starting delivery/result polling cycle")
    registry = Registry(settings.WORK_DIR)
    outputs = []
    sender = (
        (settings.N8N_RESPONSE_SENDER or settings.INTERNAL_OUTLOOK_EMAIL)
        .lower()
        .strip()
    )
    result_settings = settings.input_mailbox()
    handled_path = settings.WORK_DIR / "result-mail-handled.json"
    if not sender:
        raise ValueError("Response sender is required")
    with locked(settings.WORK_DIR / "result-agent.lock"):
        handled = _handled_response_messages(handled_path)
        for uid, uidvalidity, message in _iter_messages(
            settings, None, unseen_only=False
        ):
            token = f"{uidvalidity}:{uid.decode()}"
            if token in handled:
                continue
            if parseaddr(str(message.get("From", "")))[1].lower() != sender:
                logger.warning("Ignoring unexpected sender, UID %s", uid.decode())
                _remember_response(handled_path, handled, token)
                continue
            try:
                attachments = [
                    p
                    for p in message.iter_attachments()
                    if (p.get_filename() or "").lower().endswith(".json")
                ]
                if len(attachments) != 1:
                    raise ValueError(
                        "Expected exactly one JSON attachment per response"
                    )
                payload = json.loads(
                    attachments[0].get_payload(decode=True).decode("utf-8-sig")
                )
                key = correlate(str(message.get("Subject", "")), payload)
                record = registry.get(key)
                if not record:
                    raise ValueError("Unknown response UID")
                if not record.get("outlook_received_at"):
                    registry.put(
                        {
                            "request_id": key,
                            "outlook_received_at": str(
                                message.get("X-Agent-Received-At", "")
                            ),
                        }
                    )
                if record["status"] == "completed":
                    mark_processed(settings, uid)
                    _remember_response(handled_path, handled, token)
                    continue
                if record["status"] != "sent":
                    # This can be a response that raced ahead of the local CSV
                    # transition. Keep it retryable instead of blacklisting it.
                    raise RuntimeError("Request submission is not yet confirmed")
                recipient = record.get("sender", "").strip()
                if not re.fullmatch(r"[^\s@,;<>]+@[^\s@,;<>]+", recipient):
                    raise ValueError(
                        "Missing or invalid original sender in CSV; result not sent"
                    )
                directory = settings.WORK_DIR / "jobs" / key
                original_name = record["filename"].replace("\\", "/").split("/")[-1]
                safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", original_name).rstrip(
                    " ."
                )
                output = directory / ("out_" + (safe_name or "request.xlsx"))
                submission = directory / "result-submission.json"
                if not submission.exists():
                    forecasts = forecast_rows(payload)
                    fill_forecasts(
                        directory / "source.xlsx", output, forecasts, settings
                    )
                    registry.put(
                        {
                            "request_id": key,
                            "output_filename": output.name,
                            "output_path": str(output.resolve()),
                        }
                    )
                    write_json(directory / "response.json", {"payload": payload})
                    submit_once(
                        submission,
                        lambda recipient=recipient, key=key, output=output: (
                            send_result_workbook(
                                result_settings, recipient, key, output
                            )
                        ),
                    )
                else:
                    # An uncertain send blocks retries; an accepted one is not repeated.
                    submit_once(submission, lambda: None)
                registry.put(
                    {
                        "request_id": key,
                        "status": "completed",
                        "output_path": str(output.resolve()),
                        "output_filename": output.name,
                        "user_sent_at": json.loads(
                            submission.read_text(encoding="utf-8")
                        ).get("accepted_at", ""),
                        "user_message_id": json.loads(
                            submission.read_text(encoding="utf-8")
                        ).get("message_id", ""),
                    }
                )
                logger.info(
                    "Result %s sent to original requester; output=%s", key, output.name
                )
                mark_processed(settings, uid)
                outputs.append(output)
                _remember_response(handled_path, handled, token)
                # The second agent completes at most one queue item per cycle.
                return outputs
            except (ValueError, KeyError, TypeError):
                logger.exception(
                    "Response processing failed for mail UID %s", uid.decode()
                )
                # Email content is immutable. A corrected n8n response will have a
                # new IMAP UID, so do not log the same terminal validation error
                # every polling cycle.
                _remember_response(handled_path, handled, token)
            except Exception:
                # Transport, filesystem and uncertain SMTP errors remain retryable.
                logger.exception(
                    "Response processing failed for mail UID %s; will retry",
                    uid.decode(),
                )
        # No actionable response was completed: submit one prepared request.
        record = registry.first_with_status("prepared")
        if record:
            outputs.append(_send_prepared_request(settings, registry, record))
    return outputs


def run_agent(settings, role, once=False, stop_event=None):
    action = process_requests if role == "requests" else process_results
    while True:
        if stop_event is not None and stop_event.is_set():
            return
        try:
            for result in action(settings):
                print(result, flush=True)
        except Exception:
            if once:
                raise
            logger.exception("Mail polling failed; retrying next cycle")
        if once:
            return
        if stop_event is None:
            time.sleep(settings.POLL_INTERVAL_SECONDS)
        elif stop_event.wait(settings.POLL_INTERVAL_SECONDS):
            return
