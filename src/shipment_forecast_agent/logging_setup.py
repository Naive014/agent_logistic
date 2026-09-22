import logging
from logging.handlers import RotatingFileHandler


def configure_agent_logging(settings, role):
    directory = settings.WORK_DIR / "logs"
    directory.mkdir(parents=True, exist_ok=True)
    path = (directory / f"{role}-agent.log").resolve()
    root = logging.getLogger()
    for handler in list(root.handlers):
        if getattr(handler, "_agent_log", False):
            if handler.baseFilename == str(path):
                return
            root.removeHandler(handler)
            handler.close()
    handler = RotatingFileHandler(
        path, maxBytes=5_000_000, backupCount=5, encoding="utf-8"
    )
    handler._agent_log = True
    handler.setFormatter(
        logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s")
    )
    root.addHandler(handler)
    root.setLevel(logging.INFO)
