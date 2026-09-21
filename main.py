"""Windows PPT → iSpring Cloud worker.

Long-running outbound poller. Talks only to the PHP queue APIs:

- GET  /pptupdate/ispringcloud/next
- POST /pptupdate/ispringcloud/result

Never connects to MySQL. Never calls /retry.
"""

from __future__ import annotations

import logging
import signal

from api_client.jobs import BackendClient
from config.settings import get_settings
from utils.logging_config import configure_logging
from utils.slack import SlackNotifier
from worker.cloud_pipeline import CloudPipeline
from worker.loop import WorkerLoop

logger = logging.getLogger(__name__)


def main() -> int:
    settings = get_settings()
    configure_logging(settings)
    if not settings.backend_base_url:
        logger.error("BACKEND_BASE_URL is required")
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
