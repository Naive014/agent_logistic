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
    message = EmailMessage()
    message["From"] = "robot@example.com"
    message["To"] = "recipient@example.com"
    with caplog.at_level("INFO"):
        internet_mail._send_now(Settings(_env_file=None), message)
    client.send_message.assert_called_once_with(message)
    assert message["Message-ID"]
    assert "SMTP accepted message" in caplog.text
    if quit_fails:
        client.close.assert_called_once()


def test_rejection_is_not_reported_as_success(monkeypatch, caplog):
    client = Mock()
    client.send_message.return_value = {"recipient@example.com": (550, b"Rejected")}
    monkeypatch.setattr(internet_mail, "_open_smtp", lambda settings: client)
    with pytest.raises(smtplib.SMTPRecipientsRefused):
        internet_mail._send_now(Settings(_env_file=None), EmailMessage())
    assert "SMTP accepted message" not in caplog.text
    client.quit.assert_called_once()
