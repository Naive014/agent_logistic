from datetime import date, datetime
from typing import Any
from uuid import UUID

from pydantic import BaseModel, ConfigDict, Field, model_validator


class Direction(BaseModel):
    model_config = ConfigDict(extra="allow")

    direction_id: str
    shipping_point: str
    delivery_region: str
    attributes: dict[str, Any] = Field(default_factory=dict)


class ForecastRequest(BaseModel):
    schema_version: str = "1.0"
    request_id: UUID
    created_at: datetime
    reply_to: str
    directions: list[Direction]


class ForecastRow(BaseModel):
    direction_id: str
    shipping_point: str
    delivery_region: str
    period: date
    forecast: float = Field(ge=0, allow_inf_nan=False)
    attributes: dict[str, Any] = Field(default_factory=dict)


class ForecastResponse(BaseModel):
    schema_version: str = "1.0"
    request_id: UUID
    generated_at: datetime
    forecasts: list[ForecastRow]

    @model_validator(mode="after")
    def unique_grain(self):
        keys = [(row.direction_id, row.period) for row in self.forecasts]
        if len(keys) != len(set(keys)):
            raise ValueError(
                "Duplicate direction_id + period rows in forecast response"
            )
        return self
