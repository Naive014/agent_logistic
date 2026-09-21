from datetime import UTC, date, datetime
from uuid import uuid4

import pytest
from pydantic import ValidationError

from shipment_forecast_agent.models import ForecastResponse, ForecastRow


def test_response_rejects_duplicate_direction_period():
    row = ForecastRow(
        direction_id="DIR-0001",
        shipping_point="Малино",
        delivery_region="Смоленская область",
        period=date(2026, 9, 1),
        forecast=9,
    )
    with pytest.raises(ValidationError):
        ForecastResponse(
            request_id=uuid4(),
            generated_at=datetime.now(UTC),
            forecasts=[row, row],
        )
