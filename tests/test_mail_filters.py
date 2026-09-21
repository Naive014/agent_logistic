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


def test_plaintext_smtp_is_rejected():
    with pytest.raises(ValueError, match="TLS"):
        mail._open_smtp(
            Settings(_env_file=None, SMTP_USE_SSL=False, SMTP_USE_STARTTLS=False)
        )


def test_imap_failure_to_mark_is_visible(monkeypatch):
    client = Mock()
    client.select.return_value = ("OK", [])
    client.uid.return_value = ("NO", [])
    monkeypatch.setattr(mail, "_open_imap", lambda settings: client)
    with pytest.raises(RuntimeError, match="mark"):
        mail.mark_processed(Settings(_env_file=None), b"1")
    client.logout.assert_called_once()
