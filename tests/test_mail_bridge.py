import pytest
from openpyxl import Workbook

from shipment_forecast_agent import workflow
from shipment_forecast_agent.config import Settings
from shipment_forecast_agent.internet_mail import (
    ImapInputEnvelope,
    send_forecast_request,
)
from shipment_forecast_agent.models import ForecastRequest


def test_send_requires_internal_outlook_address():
    settings = Settings(
        _env_file=None,
        MAILBOX_NAME="robot@example.com",
        MAILBOX_PASSWORD="secret",
        INTERNAL_OUTLOOK_EMAIL="",
    )
    request = ForecastRequest.model_validate(
        {
            "request_id": "85f60ba1-01be-44b9-b079-2b2f3c42ccb6",
            "created_at": "2026-09-20T10:00:00Z",
            "reply_to": "requester@example.com",
            "directions": [],
        }
    )
    with pytest.raises(ValueError, match="INTERNAL_OUTLOOK_EMAIL"):
        send_forecast_request(settings, request)


def test_process_input_mail_forwards_request(monkeypatch, tmp_path):
    source = tmp_path / "source.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Пункт отгрузки", "Регион доставки"])
    sheet.append(["Малино", "Смоленская область"])
    workbook.save(source)

    envelope = ImapInputEnvelope(
        uid=b"42",
        sender="requester@example.com",
        subject="[FORECAST_INPUT] test",
        filename="directions.xlsx",
        content=source.read_bytes(),
    )
    sent = []
    marked = []
    monkeypatch.setattr(workflow, "iter_input_workbooks", lambda settings: [envelope])
    monkeypatch.setattr(
        workflow,
        "send_forecast_request",
        lambda settings, request: sent.append(request),
    )
    monkeypatch.setattr(
        workflow, "mark_processed", lambda settings, uid: marked.append(uid)
    )
    settings = Settings(
        _env_file=None,
        WORK_DIR=tmp_path / "workdir",
        MAILBOX_NAME="robot@example.com",
        MAILBOX_PASSWORD="secret",
        INTERNAL_OUTLOOK_EMAIL="n8n@company.ru",
    )

    requests = workflow.process_input_messages(settings)

    assert len(requests) == 1
    assert requests[0].reply_to == "requester@example.com"
    assert requests[0].directions[0].shipping_point == "Малино"
    assert sent == requests
    assert marked == [b"42"]
    assert list((tmp_path / "workdir" / "incoming").glob("*/directions.xlsx"))
