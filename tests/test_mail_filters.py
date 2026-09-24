from email.message import EmailMessage
from unittest.mock import Mock

import pytest

from shipment_forecast_agent import internet_mail as mail
from shipment_forecast_agent.config import Settings


def message(sender, filename, content):
    msg = EmailMessage()
    msg["From"] = sender
    msg["Subject"] = "[FORECAST_INPUT] test"
    msg.set_content("Test")
    msg.add_attachment(
        content, maintype="application", subtype="octet-stream", filename=filename
    )
    return msg


def test_response_sender_checked_before_parsing(monkeypatch):
    msg = message("intruder@example.com", "response.json", b"invalid")
    monkeypatch.setattr(mail, "_iter_messages", lambda *args: [(b"1", "7", msg)])
    assert (
        list(
            mail.iter_forecast_responses(
                Settings(_env_file=None, INTERNAL_OUTLOOK_EMAIL="n8n@example.com")
            )
        )
        == []
    )


def test_invalid_json_is_not_acknowledged(monkeypatch):
    msg = message("n8n@example.com", "response.json", b"invalid")
    monkeypatch.setattr(mail, "_iter_messages", lambda *args: [(b"1", "7", msg)])
    assert (
        list(
            mail.iter_forecast_responses(
                Settings(_env_file=None, INTERNAL_OUTLOOK_EMAIL="n8n@example.com")
            )
        )
        == []
    )


def test_attachment_filename_cannot_escape_directory(monkeypatch):
    msg = message("author@example.com", "../../CON.xlsx", b"workbook")
    monkeypatch.setattr(mail, "_iter_messages", lambda *args: [(b"1", "7", msg)])
    (envelope,) = mail.iter_input_workbooks(Settings(_env_file=None))
    assert envelope.filename.startswith("input-")
    assert "/" not in envelope.filename and "\\" not in envelope.filename
    assert envelope.uidvalidity == "7"


def test_input_workbook_accepts_any_subject_and_keeps_thread_headers(monkeypatch):
    msg = message("author@example.com", "request.xlsx", b"workbook")
    msg.replace_header("Subject", "Произвольная тема")
    msg["Message-ID"] = "<request@example.com>"
    msg["References"] = "<parent@example.com>"
    calls = []

    def messages(settings, prefix):
        calls.append(prefix)
        return [(b"2", "8", msg)]

    monkeypatch.setattr(mail, "_iter_messages", messages)
    (envelope,) = mail.iter_input_workbooks(Settings(_env_file=None))

    assert calls == [None]
    assert envelope.subject == "Произвольная тема"
    assert envelope.message_id == "<request@example.com>"
    assert envelope.references == "<parent@example.com>"


def test_result_is_a_reply_with_original_attachment_name(monkeypatch, tmp_path):
    workbook = tmp_path / "generated.xlsx"
    workbook.write_bytes(b"xlsx")
    sent = []
    monkeypatch.setattr(
        mail, "_send_now", lambda settings, message: sent.append(message)
    )

    mail.send_result_workbook(
        Settings(_env_file=None),
        "author@example.com",
        "RQ-" + "a" * 32,
        workbook,
        original_subject="Запрос прогноза",
        original_message_id="<request@example.com>",
        original_references="<parent@example.com>",
        attachment_filename="Запрос ТЗ.xlsx",
    )

    message = sent[0]
    assert message["Subject"] == "RE: Запрос прогноза"
    assert message["In-Reply-To"] == "<request@example.com>"
    assert message["References"] == "<parent@example.com> <request@example.com>"
    assert next(message.iter_attachments()).get_filename() == "Запрос ТЗ.xlsx"


def test_plaintext_smtp_is_rejected():
    with pytest.raises(ValueError, match="TLS"):
        mail._open_smtp(
            Settings(_env_file=None, SMTP_USE_SSL=False, SMTP_USE_STARTTLS=False)
        )


@pytest.mark.parametrize(
    "subject",
    [
        "[FORECAST_INPUT]",
        "FW: [FORECAST_INPUT]",
        "Fwd: [FORECAST_INPUT] request",
        "RE: FW: [FORECAST_INPUT]",
        "ПЕР: [FORECAST_INPUT]",
        "ОТВ: [FORECAST_INPUT]",
    ],
)
def test_input_tag_accepts_reply_and_forward_prefixes(subject):
    assert mail._subject_matches_prefix(subject, "[FORECAST_INPUT]")


@pytest.mark.parametrize(
    "subject",
    ["Other [FORECAST_INPUT]", "FW: Other [FORECAST_INPUT]", "[OTHER]"],
)
def test_input_tag_is_not_accepted_in_arbitrary_subject_text(subject):
    assert not mail._subject_matches_prefix(subject, "[FORECAST_INPUT]")


def test_imap_failure_to_mark_is_visible(monkeypatch):
    client = Mock()
    client.select.return_value = ("OK", [])
    client.uid.return_value = ("NO", [])
    monkeypatch.setattr(mail, "_open_imap", lambda settings: client)
    with pytest.raises(RuntimeError, match="mark"):
        mail.mark_processed(Settings(_env_file=None), b"1")
    client.logout.assert_called_once()
