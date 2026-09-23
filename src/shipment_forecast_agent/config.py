from functools import lru_cache
from pathlib import Path
from typing import Annotated

from pydantic import Field, field_validator
from pydantic_settings import BaseSettings, NoDecode, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", extra="ignore")

    APP_ENV: str = Field(default="dev", pattern="^(dev|prod)$")
    WORK_DIR: Path = Path("workdir")
    MAILBOX_NAME: str = ""
    MAILBOX_PASSWORD: str = Field(default="", repr=False)
    INPUT_MAILBOX_NAME: str = ""
    INPUT_MAILBOX_PASSWORD: str = Field(default="", repr=False)
    IMAP_SERVER: str = "imap.beget.com"
    IMAP_PORT: int = Field(default=993, ge=1, le=65535)
    IMAP_USE_SSL: bool = True
    SMTP_SERVER: str = "smtp.beget.com"
    SMTP_PORT: int = Field(default=465, ge=1, le=65535)
    SMTP_USE_SSL: bool = True
    SMTP_USE_STARTTLS: bool = False
    SAVE_SENT_COPY: bool = True
    IMAP_SENT_FOLDER: str = ""
    MAIL_TIMEOUT_SECONDS: int = Field(default=30, ge=1)
    POLL_INTERVAL_SECONDS: int = Field(default=30, ge=1)
    FORECAST_START_MONTH: str = ""
    FORECAST_HORIZON_MONTHS: int = Field(default=3, ge=1, le=36)
    EXCEL_SHEET: str = ""
    EXCEL_HEADER_ROW: int = Field(default=1, ge=1)
    RESULT_MAILBOX: str = ""
    N8N_RESPONSE_SENDER: str = ""
    INTERNAL_OUTLOOK_EMAIL: str = ""
    SEND_RESULTS_TO_REQUESTER: bool = False
    INPUT_SUBJECT_PREFIX: str = "[FORECAST_INPUT]"
    N8N_REQUEST_SUBJECT_PREFIX: str = "[N8N_FORECAST_REQUEST]"
    N8N_RESPONSE_SUBJECT_PREFIX: str = "[N8N_FORECAST_RESULT]"
    TRUSTED_REQUESTERS: Annotated[tuple[str, ...], NoDecode] = ()

    def input_mailbox(self) -> "Settings":
        """Input/result mailbox; empty credentials retain single-mailbox mode."""
        if not self.INPUT_MAILBOX_NAME and not self.INPUT_MAILBOX_PASSWORD:
            return self
        if not self.INPUT_MAILBOX_NAME or not self.INPUT_MAILBOX_PASSWORD:
            raise ValueError(
                "Both INPUT_MAILBOX_NAME and INPUT_MAILBOX_PASSWORD are required"
            )
        return self.model_copy(
            update={
                "MAILBOX_NAME": self.INPUT_MAILBOX_NAME,
                "MAILBOX_PASSWORD": self.INPUT_MAILBOX_PASSWORD,
            }
        )

    @field_validator("TRUSTED_REQUESTERS", mode="before")
    @classmethod
    def parse_emails(cls, value):
        if isinstance(value, str):
            return tuple(x.strip().lower() for x in value.split(",") if x.strip())
        return value


@lru_cache
def get_settings() -> Settings:
    return Settings()
