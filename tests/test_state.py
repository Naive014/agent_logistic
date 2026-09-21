import json
from unittest.mock import Mock

import pytest

from shipment_forecast_agent.state import UncertainSubmission, submit_once


def test_accepted_submission_is_not_repeated(tmp_path):
    operation = Mock()
    path = tmp_path / "submission.json"
    assert submit_once(path, operation)
    assert not submit_once(path, operation)
    operation.assert_called_once()
    assert json.loads(path.read_text())["status"] == "smtp_accepted"


def test_failed_submission_requires_manual_check(tmp_path):
    operation = Mock(side_effect=OSError("Connection lost"))
    path = tmp_path / "submission.json"
    with pytest.raises(OSError):
        submit_once(path, operation)
    with pytest.raises(UncertainSubmission):
        submit_once(path, operation)
    operation.assert_called_once()


def test_incomplete_journal_never_resends(tmp_path):
    path = tmp_path / "submission.json"
    path.write_text("")
    operation = Mock()
    with pytest.raises(UncertainSubmission):
        submit_once(path, operation)
    operation.assert_not_called()
