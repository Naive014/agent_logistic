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
from .template_excel import current_month, fill_forecasts, request_data, validate_tender

logger = logging.getLogger(__name__)
UID_PATTERN = r"RQ-[a-f0-9]{32}"


def process_requests(settings):
    configure_agent_logging(settings, "request")
    logger.info("Starting request polling cycle")
    registry = Registry(settings.WORK_DIR)
    incoming = settings.input_mailbox()
    processed, failed, acknowledged = [], set(), set()
    with locked(settings.WORK_DIR / "request-agent.lock"):
        for item in iter_input_workbooks(incoming):
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
                if record and record["status"] in ("sent", "completed"):
                    acknowledged.add(item.uid)
                    continue
                if not source.exists():
                    source.write_bytes(item.content)
                payload_path = directory / "request.json"
                if payload_path.exists():
                    payload = json.loads(payload_path.read_text(encoding="utf-8"))
                    if not (directory / "request-submission.json").exists():
                        validate_tender(source, settings)
                        payload.pop("uid", None)
                        payload["request_id"] = key
                        payload.setdefault("ppr_month", current_month())
                        payload["mail_id"] = key
                        write_json(payload_path, payload)
                else:
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
                if not settings.INTERNAL_OUTLOOK_EMAIL:
                    raise ValueError("INTERNAL_OUTLOOK_EMAIL is required")

                def send():
                    message = EmailMessage()
                    message["From"] = settings.MAILBOX_NAME
                    message["To"] = settings.INTERNAL_OUTLOOK_EMAIL
                    message["Reply-To"] = settings.MAILBOX_NAME
                    message["Subject"] = f"{settings.N8N_REQUEST_SUBJECT_PREFIX} {key}"
                    message["Date"] = format_datetime(datetime.now(UTC))
                    message.set_content(
                        f"Запрос прогноза. Верните mail_id {key} в JSON ответа."
                    )
                    message.add_attachment(
                        json.dumps(payload, ensure_ascii=False).encode("utf-8"),
                        maintype="application",
                        subtype="json",
                        filename="forecast_request.json",
                    )
                    return _send_now(settings, message)

                submit_once(directory / "request-submission.json", send)
                receipt = json.loads(
                    (directory / "request-submission.json").read_text(encoding="utf-8")
                )
                registry.put(
                    {
                        "request_id": key,
                        "status": "sent",
                        "outlook_sent_at": receipt.get("accepted_at", ""),
                    }
                )
                logger.info("Request %s accepted by SMTP", key)
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
        raise ValueError("Invalid JSON shape")
    if (
        not isinstance(result, list)
        or not result
        or not all(isinstance(r, dict) for r in result)
    ):
        raise ValueError("Response contains no forecast rows")
    return result


def process_results(settings):
    configure_agent_logging(settings, "result")
    logger.info("Starting result polling cycle")
    registry = Registry(settings.WORK_DIR)
    outputs = []
    sender = (
        (settings.N8N_RESPONSE_SENDER or settings.INTERNAL_OUTLOOK_EMAIL)
        .lower()
        .strip()
    )
    result_settings = settings.input_mailbox()
    if not sender:
        raise ValueError("Response sender is required")
    with locked(settings.WORK_DIR / "result-agent.lock"):
        for uid, _, message in _iter_messages(
            settings, settings.N8N_RESPONSE_SUBJECT_PREFIX
        ):
            if parseaddr(str(message.get("From", "")))[1].lower() != sender:
                logger.warning("Ignoring unexpected sender, UID %s", uid.decode())
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
                    continue
                if record["status"] != "sent":
                    raise ValueError("Request submission is not yet confirmed")
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
                        lambda: send_result_workbook(
                            result_settings, recipient, key, output
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
            except Exception:
                logger.exception(
                    "Response processing failed for mail UID %s", uid.decode()
                )
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
