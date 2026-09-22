"""Durable submission journal: SMTP acceptance is not recipient delivery."""

import json
import os
from datetime import UTC, datetime
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Callable


class UncertainSubmission(RuntimeError):
    """An interrupted submission requires inspection, never an automatic resend."""


def write_json(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(
        mode="w", encoding="utf-8", dir=path.parent, delete=False
    ) as handle:
        temporary = Path(handle.name)
        json.dump(value, handle, ensure_ascii=False, indent=2)
        handle.flush()
        os.fsync(handle.fileno())
    try:
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


def submit_once(path: Path, operation: Callable[[], str | None]) -> bool:
    """Claim before sending; return False for a previously accepted submission.

    Exclusive creation protects concurrent runs. A crash or transport failure leaves
    a submitting record; an operator must check server logs before clearing it.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("x", encoding="utf-8") as handle:
            json.dump({"status": "submitting"}, handle)
            handle.flush()
            os.fsync(handle.fileno())
    except FileExistsError:
        try:
            status = json.loads(path.read_text(encoding="utf-8")).get("status")
        except (ValueError, OSError):
            status = None
        if status == "smtp_accepted":
            return False
        raise UncertainSubmission(
            f"Check SMTP logs before retrying submission: {path}"
        ) from None
    message_id = operation()
    write_json(
        path,
        {
            "status": "smtp_accepted",
            "accepted_at": datetime.now(UTC).isoformat(),
            "message_id": message_id if isinstance(message_id, str) else "",
        },
    )
    return True
