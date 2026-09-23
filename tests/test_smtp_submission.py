import smtplib
from email.message import EmailMessage
from unittest.mock import Mock

import pytest

from shipment_forecast_agent import internet_mail
from shipment_forecast_agent.config import Settings


@pytest.mark.parametrize("quit_fails", [False, True])
def test_direct_submission_remains_successful_after_acceptance(
    monkeypatch, caplog, quit_fails
):
    client = Mock()
    client.send_message.return_value = {}
    if quit_fails:
        client.quit.side_effect = smtplib.SMTPServerDisconnected("closed")
    monkeypatch.setattr(internet_mail, "_open_smtp", lambda settings: client)
    saved = []
    monkeypatch.setattr(
        internet_mail,
        "_append_sent_copy",
        lambda settings, message: saved.append(message),
    )
    message = EmailMessage()
    message["From"] = "robot@example.com"
    message["To"] = "recipient@example.com"
    with caplog.at_level("INFO"):
        internet_mail._send_now(Settings(_env_file=None), message)
    client.send_message.assert_called_once_with(message)
    assert message["Message-ID"]
    assert saved == [message]
    assert "SMTP accepted message" in caplog.text
    if quit_fails:
        client.close.assert_called_once()


def test_rejection_is_not_reported_as_success(monkeypatch, caplog):
    client = Mock()
    client.send_message.return_value = {"recipient@example.com": (550, b"Rejected")}
    monkeypatch.setattr(internet_mail, "_open_smtp", lambda settings: client)
    saved = []
    monkeypatch.setattr(
        internet_mail,
        "_append_sent_copy",
        lambda settings, message: saved.append(message),
    )
    with pytest.raises(smtplib.SMTPRecipientsRefused):
        internet_mail._send_now(Settings(_env_file=None), EmailMessage())
    assert "SMTP accepted message" not in caplog.text
    assert saved == []
    client.quit.assert_called_once()


def test_sent_copy_failure_never_retries_accepted_message(monkeypatch, caplog):
    smtp = Mock()
    smtp.send_message.return_value = {}
    monkeypatch.setattr(internet_mail, "_open_smtp", lambda settings: smtp)
    monkeypatch.setattr(
        internet_mail,
        "_append_sent_copy",
        lambda settings, message: (_ for _ in ()).throw(RuntimeError("IMAP down")),
    )
    message = EmailMessage()
    message["From"] = "robot@example.com"
    message["To"] = "recipient@example.com"

    with caplog.at_level("ERROR"):
        message_id = internet_mail._send_now(Settings(_env_file=None), message)

    assert message_id == message["Message-ID"]
    smtp.send_message.assert_called_once_with(message)
    assert "was sent, but its IMAP Sent copy could not be saved" in caplog.text


def test_sent_folder_is_detected_from_imap_special_use_flag():
    client = Mock()
    client.list.return_value = (
        "OK",
        [
            b'(\\HasNoChildren) "/" "INBOX"',
            b'(\\HasNoChildren \\Sent) "/" "Sent Items"',
        ],
    )

    assert internet_mail._sent_mailbox(client) == b'"Sent Items"'


def test_configured_sent_folder_is_quoted():
    assert internet_mail._sent_mailbox(Mock(), "Sent Items") == b'"Sent Items"'
