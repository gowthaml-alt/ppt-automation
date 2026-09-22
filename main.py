"""Windows PPT → iSpring Cloud worker.

Long-running outbound poller. Talks only to the PHP queue APIs:

- GET  /pptupdate/ispringcloud/next
- POST /pptupdate/ispringcloud/result

Never connects to MySQL. Never calls /retry.
"""

from __future__ import annotations

import logging
import signal
from pathlib import Path

from api_client.jobs import BackendClient
from config.settings import get_settings
from utils.logging_config import configure_logging
from utils.slack import SlackNotifier
from worker.cloud_pipeline import CloudPipeline
from worker.loop import WorkerLoop

logger = logging.getLogger(__name__)


def writable(folder: str, name: str) -> str:
    """Can the worker actually write here? Returns "" when it can."""
    if not folder:
        return f"{name} is not set"
    path = Path(folder)
    try:
        path.mkdir(parents=True, exist_ok=True)
        probe = path / ".write-test"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
    except OSError as exc:
        return f"{name}={folder} cannot be written to ({exc.strerror or exc})"
    return ""


def main() -> int:
    settings = get_settings()
    configure_logging(settings)
    if not settings.backend_base_url:
        logger.error("BACKEND_BASE_URL is required")
        return 2

    # Refuse to start rather than take jobs we cannot do.
    #
    # TEMP_ROOT is where each deck is downloaded, so an unwritable one fails
    # every job. Worse, LOG_ROOT holds the health file that is supposed to
    # stop a broken worker after three failures - with both unset the worker
    # would empty the whole queue into the failed pile in twenty minutes,
    # which is the exact thing that file exists to prevent. Better to stop
    # here, with the folder named, than to mark a hundred decks failed.
    problems = [
        message
        for message in (
            writable(settings.temp_root, "TEMP_ROOT"),
            writable(settings.log_root, "LOG_ROOT"),
        )
        if message
    ]
    if problems:
        for message in problems:
            logger.error(message)
        logger.error("fix these in .env and start again; no jobs were taken")
        return 2

    backend = BackendClient(settings)
    pipeline = CloudPipeline(settings, backend, slack=SlackNotifier(settings))
    loop = WorkerLoop(settings, backend, pipeline)

    def _handle_stop(signum, frame):
        logger.info("received stop signal %s", signum)
        loop.stop()

    signal.signal(signal.SIGINT, _handle_stop)
    if hasattr(signal, "SIGTERM"):
        signal.signal(signal.SIGTERM, _handle_stop)

    logger.info(
        "worker starting",
        extra={
            "app_env": settings.app_env,
            "ispring_adapter": settings.ispring_adapter,
        },
    )
    loop.run_forever()
    logger.info("worker stopped")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
