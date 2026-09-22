"""CSV shared by the two agents, protected by a short OS-level lock."""

import csv
import os
from contextlib import contextmanager
from pathlib import Path
from tempfile import NamedTemporaryFile

FIELDS = [
    "UID",
    "request_received_at",
    "filename",
    "outlook_sent_at",
    "outlook_received_at",
    "output_filename",
    "user_sent_at",
    "user_message_id",
    "request_id",
    "mailbox",
    "uidvalidity",
    "mail_uid",
    "source_path",
    "sender",
    "status",
    "output_path",
]


@contextmanager
def locked(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a+b") as handle:
        if handle.seek(0, 2) == 0:
            handle.write(b"0")
            handle.flush()
        handle.seek(0)
        if os.name == "nt":
            import msvcrt

            msvcrt.locking(handle.fileno(), msvcrt.LK_LOCK, 1)
        else:
            import fcntl

            fcntl.flock(handle, fcntl.LOCK_EX)
        try:
            yield
        finally:
            handle.seek(0)
            if os.name == "nt":
                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(handle, fcntl.LOCK_UN)


class Registry:
    def __init__(self, root: Path):
        self.path = root / "requests.csv"

    def _read(self):
        if not self.path.exists():
            return []
        with self.path.open(encoding="utf-8-sig", newline="") as handle:
            records = list(csv.DictReader(handle))
        for record in records:
            record.setdefault("UID", record.get("request_id", ""))
            for field in FIELDS:
                record.setdefault(field, "")
        return records

    def get(self, request_id):
        with locked(self.path.with_suffix(".lock")):
            return next(
                (r for r in self._read() if r["request_id"] == request_id), None
            )

    def put(self, record):
        with locked(self.path.with_suffix(".lock")):
            records = self._read()
            old = next(
                (r for r in records if r["request_id"] == record["request_id"]), None
            )
            if old is None:
                old = dict.fromkeys(FIELDS, "")
                records.append(old)
            old.update(record)
            old["UID"] = old["request_id"]
            # Preserve the original CSV before the first schema upgrade.
            if self.path.exists():
                with self.path.open(encoding="utf-8-sig", newline="") as source:
                    header = next(csv.reader(source), [])
                if "request_received_at" not in header:
                    import shutil

                    backup = self.path.with_suffix(".legacy.csv")
                    if not backup.exists():
                        shutil.copy2(self.path, backup)
            with NamedTemporaryFile(
                mode="w",
                encoding="utf-8-sig",
                newline="",
                dir=self.path.parent,
                delete=False,
            ) as handle:
                temp = Path(handle.name)
                writer = csv.DictWriter(handle, fieldnames=FIELDS)
                writer.writeheader()
                writer.writerows(records)
                handle.flush()
                os.fsync(handle.fileno())
            try:
                temp.replace(self.path)
            finally:
                temp.unlink(missing_ok=True)
