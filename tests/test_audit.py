import csv
import json
import logging

from shipment_forecast_agent.config import Settings
from shipment_forecast_agent.logging_setup import configure_agent_logging
from shipment_forecast_agent.registry import FIELDS, Registry
from shipment_forecast_agent.state import submit_once


def test_legacy_csv_is_preserved_and_upgraded(tmp_path):
    path = tmp_path / "requests.csv"
    original = (
        "request_id,filename,status,sender\nRQ-old,old.xlsx,sent,author@example.com\n"
    )
    path.write_text(original, encoding="utf-8")
    registry = Registry(tmp_path)
    registry.put({"request_id": "RQ-new", "filename": "new.xlsx", "status": "prepared"})
    assert path.with_suffix(".legacy.csv").read_text(encoding="utf-8") == original
    with path.open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        rows = list(reader)
        assert reader.fieldnames[:8] == FIELDS[:8]
    assert rows[0]["UID"] == "RQ-old"
    assert rows[0]["request_received_at"] == ""
    assert rows[0]["status"] == "sent"
    assert len(rows) == 2


def test_smtp_receipt_keeps_message_id_and_timestamp(tmp_path):
    path = tmp_path / "submission.json"
    submit_once(path, lambda: "<test@example.com>")
    receipt = json.loads(path.read_text())
    assert receipt["message_id"] == "<test@example.com>"
    assert receipt["accepted_at"].endswith("+00:00")
    assert not submit_once(path, lambda: "<duplicate@example.com>")
    assert json.loads(path.read_text()) == receipt


def test_separate_agent_logs_and_no_duplicate_handlers(tmp_path):
    settings = Settings(_env_file=None, WORK_DIR=tmp_path)
    log = logging.getLogger("test.audit")
    try:
        configure_agent_logging(settings, "request")
        configure_agent_logging(settings, "request")
        log.info("request-only-marker")
        configure_agent_logging(settings, "result")
        log.error("result-only-marker")
    finally:
        root = logging.getLogger()
        for handler in list(root.handlers):
            if getattr(handler, "_agent_log", False):
                handler.flush()
                handler.close()
                root.removeHandler(handler)
    requests = (tmp_path / "logs/request-agent.log").read_text(encoding="utf-8")
    results = (tmp_path / "logs/result-agent.log").read_text(encoding="utf-8")
    assert requests.count("request-only-marker") == 1
    assert "result-only-marker" not in requests
    assert "result-only-marker" in results
    assert "request-only-marker" not in results
