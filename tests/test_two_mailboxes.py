import json
from datetime import UTC, datetime

import pytest
from openpyxl import Workbook

from shipment_forecast_agent import internet_mail, workflow
from shipment_forecast_agent.config import Settings
from shipment_forecast_agent.models import ForecastResponse


def test_two_mailbox_roundtrip(monkeypatch, tmp_path):
    settings = Settings(
        _env_file=None,
        WORK_DIR=tmp_path / "state",
        MAILBOX_NAME="bridge@example.com",
        MAILBOX_PASSWORD="bridge-secret",
        INPUT_MAILBOX_NAME="planning@example.com",
        INPUT_MAILBOX_PASSWORD="planning-secret",
        INTERNAL_OUTLOOK_EMAIL="corporate@example.com",
        SEND_RESULTS_TO_REQUESTER=True,
    )
    book = Workbook()
    book.active.append(["Пункт отгрузки", "Регион доставки"])
    book.active.append(["Завод", "Область"])
    path = tmp_path / "input.xlsx"
    book.save(path)

    def inputs(config):
        assert config.MAILBOX_NAME == "planning@example.com"
        assert config.MAILBOX_PASSWORD == "planning-secret"
        return [
            internet_mail.ImapInputEnvelope(
                b"42",
                "author@example.com",
                "[FORECAST_INPUT]",
                "input.xlsx",
                path.read_bytes(),
            )
        ]

    submitted = []
    marked = []
    monkeypatch.setattr(workflow, "iter_input_workbooks", inputs)
    monkeypatch.setattr(
        internet_mail,
        "_send_now",
        lambda config, message: submitted.append((config.MAILBOX_NAME, message)),
    )
    monkeypatch.setattr(
        workflow,
        "mark_processed",
        lambda config, uid: marked.append((config.MAILBOX_NAME, uid)),
    )
    (request,) = workflow.process_input_messages(settings)
    account, message = submitted[0]
    assert account == message["From"] == "bridge@example.com"
    assert message["To"] == "corporate@example.com"
    assert message["Reply-To"] == "bridge@example.com"
    payload = json.loads(next(message.iter_attachments()).get_payload(decode=True))
    assert payload["reply_to"] == "bridge@example.com"
    assert request.reply_to == "author@example.com"
    direction = request.directions[0]
    response = ForecastResponse(
        request_id=request.request_id,
        generated_at=datetime.now(UTC),
        forecasts=[
            {
                "direction_id": direction.direction_id,
                "shipping_point": direction.shipping_point,
                "delivery_region": direction.delivery_region,
                "period": "2026-10-01",
                "forecast": 100,
            }
        ],
    )

    def responses(config):
        assert config.MAILBOX_NAME == "bridge@example.com"
        return [internet_mail.ImapResponseEnvelope(b"99", response)]

    monkeypatch.setattr(workflow, "iter_forecast_responses", responses)
    outputs = workflow.receive_responses(settings)
    assert len(outputs) == 1 and outputs[0].is_file()
    account, message = submitted[1]
    assert account == message["From"] == "planning@example.com"
    assert message["To"] == "author@example.com"
    assert next(message.iter_attachments()).get_filename() == "forecast_result.xlsx"
    assert marked == [("planning@example.com", b"42"), ("bridge@example.com", b"99")]
    assert workflow.process_input_messages(settings) == []
    assert workflow.receive_responses(settings) == []
    assert len(submitted) == 2


def test_single_mailbox_compatibility():
    settings = Settings(_env_file=None)
    assert settings.input_mailbox() is settings


@pytest.mark.parametrize(
    "values",
    [
        {"INPUT_MAILBOX_NAME": "planning@example.com"},
        {"INPUT_MAILBOX_PASSWORD": "secret"},
    ],
)
def test_incomplete_input_credentials_fail(values):
    with pytest.raises(ValueError, match="Both INPUT_MAILBOX"):
        Settings(_env_file=None, **values).input_mailbox()
