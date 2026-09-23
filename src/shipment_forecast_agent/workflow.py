import hashlib
import logging
from datetime import UTC, datetime
from pathlib import Path
from uuid import NAMESPACE_URL, uuid4, uuid5

from .config import Settings
from .excel_io import read_directions, write_forecast
from .internet_mail import (
    iter_forecast_responses,
    iter_input_workbooks,
    mark_processed,
    send_forecast_request,
    send_result_workbook,
)
from .models import ForecastRequest, ForecastResponse
from .state import submit_once, write_json

logger = logging.getLogger(__name__)


def build_request(input_xlsx: str | Path, reply_to: str) -> ForecastRequest:
    return ForecastRequest(
        request_id=uuid4(),
        created_at=datetime.now(UTC),
        reply_to=reply_to,
        directions=read_directions(input_xlsx),
    )


def prepare_request(
    input_xlsx: str | Path, request_json: str | Path, reply_to: str
) -> ForecastRequest:
    request = build_request(input_xlsx, reply_to)
    output = Path(request_json)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(request.model_dump_json(indent=2), encoding="utf-8")
    return request


def convert_response(response_json: str | Path, output_xlsx: str | Path) -> Path:
    response = ForecastResponse.model_validate_json(
        Path(response_json).read_text(encoding="utf-8")
    )
    return write_forecast(output_xlsx, response)


def send_request(
    input_xlsx: str | Path,
    settings: Settings,
    reply_to: str | None = None,
) -> ForecastRequest:
    request = build_request(input_xlsx, reply_to or settings.MAILBOX_NAME)
    request_dir = settings.WORK_DIR / "requests" / str(request.request_id)
    request_dir.mkdir(parents=True, exist_ok=True)
    write_json(request_dir / "forecast_request.json", request.model_dump(mode="json"))
    submit_once(
        request_dir / "submission.json",
        lambda: send_forecast_request(settings, request),
    )
    return request


def process_input_messages(settings: Settings) -> list[ForecastRequest]:
    """Read LogRocket input mail and forward normalized JSON to corporate Outlook."""
    requests: list[ForecastRequest] = []
    processed_uids: set[bytes] = set()
    input_settings = settings.input_mailbox()
    for envelope in iter_input_workbooks(input_settings):
        # Include account and UIDVALIDITY: an IMAP UID alone is not globally unique.
        identity = "|".join(
            (
                input_settings.MAILBOX_NAME.lower(),
                envelope.uidvalidity,
                envelope.uid.decode("ascii"),
                envelope.filename,
                hashlib.sha256(envelope.content).hexdigest(),
            )
        )
        request_id = uuid5(NAMESPACE_URL, identity)
        incoming_dir = settings.WORK_DIR / "incoming" / str(request_id)
        incoming_dir.mkdir(parents=True, exist_ok=True)
        input_path = incoming_dir / envelope.filename
        input_path.write_bytes(envelope.content)
        request_dir = settings.WORK_DIR / "requests" / str(request_id)
        request_path = request_dir / "forecast_request.json"
        if request_path.exists():
            request = ForecastRequest.model_validate_json(
                request_path.read_text(encoding="utf-8")
            )
        else:
            request = build_request(input_path, envelope.sender).model_copy(
                update={"request_id": request_id}
            )
            write_json(request_path, request.model_dump(mode="json"))
        if submit_once(
            request_dir / "submission.json",
            lambda request=request: send_forecast_request(settings, request),
        ):
            requests.append(request)
        processed_uids.add(envelope.uid)
    for uid in processed_uids:
        mark_processed(input_settings, uid)
    return requests


def receive_responses(settings: Settings) -> list[Path]:
    result_settings = settings.input_mailbox()
    outputs: list[Path] = []
    for envelope in iter_forecast_responses(settings):
        request_path = (
            settings.WORK_DIR
            / "requests"
            / str(envelope.response.request_id)
            / "forecast_request.json"
        )
        if not request_path.is_file():
            logger.warning(
                "Ignored unknown response request_id=%s", envelope.response.request_id
            )
            continue
        request = ForecastRequest.model_validate_json(
            request_path.read_text(encoding="utf-8")
        )
        expected = {
            (
                direction.direction_id,
                direction.shipping_point,
                direction.delivery_region,
            )
            for direction in request.directions
        }
        actual = {
            (row.direction_id, row.shipping_point, row.delivery_region)
            for row in envelope.response.forecasts
        }
        if not actual or actual != expected:
            logger.warning(
                "Ignored incomplete or mismatched directions request_id=%s",
                request.request_id,
            )
            continue
        response_dir = (
            settings.WORK_DIR / "responses" / str(envelope.response.request_id)
        )
        response_dir.mkdir(parents=True, exist_ok=True)
        json_path = response_dir / "forecast_response.json"
        result_submission = response_dir / "submission.json"
        if result_submission.exists():
            # Never replace the delivered workbook with a different duplicate response.
            submit_once(result_submission, lambda: None)
            mark_processed(settings, envelope.uid)
            continue
        write_json(json_path, envelope.response.model_dump(mode="json"))
        output_path = write_forecast(
            response_dir / "forecast_result.xlsx", envelope.response
        )
        outputs.append(output_path)
        if settings.SEND_RESULTS_TO_REQUESTER and request.reply_to:
            submit_once(
                result_submission,
                lambda request=request, envelope=envelope, output_path=output_path: (
                    send_result_workbook(
                        result_settings,
                        request.reply_to,
                        str(envelope.response.request_id),
                        output_path,
                    )
                ),
            )
        mark_processed(settings, envelope.uid)
    return outputs
