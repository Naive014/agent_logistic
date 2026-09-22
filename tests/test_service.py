from unittest.mock import Mock

import pytest

from shipment_forecast_agent import agents, service
from shipment_forecast_agent.config import Settings


def settings():
    return Settings(
        _env_file=None,
        INTERNAL_OUTLOOK_EMAIL="corp@example.com",
        RESULT_MAILBOX="planning@example.com",
    )


def test_service_starts_and_stops_both_workers(monkeypatch):
    context = Mock()
    processes = [Mock(name="requests"), Mock(name="results")]
    for process in processes:
        process.exitcode = None
        process.is_alive.return_value = False
    context.Process.side_effect = processes
    monkeypatch.setattr(service.multiprocessing, "get_context", lambda mode: context)
    monkeypatch.setattr(service.time, "sleep", Mock(side_effect=KeyboardInterrupt))
    service.run_service(settings())
    assert [c.kwargs["args"][1] for c in context.Process.call_args_list] == [
        "requests",
        "results",
    ]
    context.Event.return_value.set.assert_called_once()
    for process in processes:
        process.start.assert_called_once()
        process.join.assert_called_once()
        process.close.assert_called_once()
        process.terminate.assert_not_called()


def test_worker_crash_stops_remaining_worker(monkeypatch):
    context = Mock()
    processes = [Mock(), Mock()]
    processes[0].exitcode = 1
    processes[0].name = "requests"
    processes[0].is_alive.return_value = False
    processes[1].exitcode = None
    processes[1].is_alive.side_effect = [True, False]
    context.Process.side_effect = processes
    monkeypatch.setattr(service.multiprocessing, "get_context", lambda mode: context)
    with pytest.raises(RuntimeError, match="exited"):
        service.run_service(settings())
    context.Event.return_value.set.assert_called_once()
    processes[1].terminate.assert_called_once()


def test_poll_wait_is_interruptible(monkeypatch):
    action = Mock(return_value=[])
    stop = Mock()
    stop.is_set.return_value = False
    stop.wait.return_value = True
    monkeypatch.setattr(agents, "process_requests", action)
    agents.run_agent(settings(), "requests", stop_event=stop)
    action.assert_called_once()
    stop.wait.assert_called_once_with(30)
