"""One foreground entry point supervising two independent mail workers."""

import logging
import multiprocessing
import signal
import time

from .agents import run_agent

logger = logging.getLogger(__name__)


def _worker(settings, role, stop_event):
    # Only the parent handles Ctrl+C, allowing workers to finish a mail operation.
    signal.signal(signal.SIGINT, signal.SIG_IGN)
    logging.basicConfig(
        level=logging.INFO,
        format=f"%(asctime)s [{role}] %(levelname)s: %(message)s",
    )
    run_agent(settings, role, stop_event=stop_event)


def run_service(settings):
    settings.input_mailbox()  # Validate the two-account configuration before spawning.
    if not settings.INTERNAL_OUTLOOK_EMAIL:
        raise ValueError("INTERNAL_OUTLOOK_EMAIL is required")
    context = multiprocessing.get_context("spawn")
    stop_event = context.Event()
    processes = []
    try:
        for role in ("requests", "results"):
            process = context.Process(
                target=_worker, args=(settings, role, stop_event), name=f"agent-{role}"
            )
            process.start()
            processes.append(process)
        logger.info("Both agents started. Press Ctrl+C to stop.")
        while True:
            for process in processes:
                if process.exitcode is not None:
                    raise RuntimeError(
                        f"{process.name} exited with code {process.exitcode}; stopping both agents"
                    )
            time.sleep(0.5)
    except KeyboardInterrupt:
        logger.info("Stopping both agents after current operations...")
    finally:
        stop_event.set()
        deadline = time.monotonic() + settings.MAIL_TIMEOUT_SECONDS + 5
        for process in processes:
            process.join(timeout=max(0, deadline - time.monotonic()))
        for process in processes:
            if process.is_alive():
                logger.warning(
                    "Force-stopping %s; inspect submitting journals before retrying uncertain sends",
                    process.name,
                )
                process.terminate()
                process.join(timeout=5)
                if process.is_alive():
                    process.kill()
                    process.join()
            process.close()
