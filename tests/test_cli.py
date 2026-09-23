import sys

import pytest

from shipment_forecast_agent import cli


@pytest.mark.parametrize(
    "command,role",
    [("request-agent", "requests"), ("result-agent", "results")],
)
def test_individual_agents_run_once_by_default(monkeypatch, command, role):
    calls = []
    settings = object()
    monkeypatch.setattr(cli, "get_settings", lambda: settings)
    monkeypatch.setattr(
        cli, "run_agent", lambda *args, **kwargs: calls.append((args, kwargs))
    )
    monkeypatch.setattr(sys, "argv", ["shipment-forecast-agent", command])

    cli.main()

    assert calls == [((settings, role), {"once": True})]


@pytest.mark.parametrize(
    "command,role",
    [("request-agent", "requests"), ("result-agent", "results")],
)
def test_continuous_flag_enables_polling_loop(monkeypatch, command, role):
    calls = []
    settings = object()
    monkeypatch.setattr(cli, "get_settings", lambda: settings)
    monkeypatch.setattr(
        cli, "run_agent", lambda *args, **kwargs: calls.append((args, kwargs))
    )
    monkeypatch.setattr(
        sys, "argv", ["shipment-forecast-agent", command, "--continuous"]
    )

    cli.main()

    assert calls == [((settings, role), {"once": False})]
