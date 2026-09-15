"""Windows PPT automation worker.

Long-running outbound poller. Does not serve HTTP jobs and does not connect
to MySQL.
"""

from __future__ import annotations

import logging
import signal
import sys

from api_client.jobs import BackendClient
from browser_test.playwright_check import NoOpBrowserChecker, PlaywrightChecker
from config.settings import Settings, get_settings
from powerpoint.service import NoOpPowerPointService, PowerPointService
from publisher.ispring import get_publisher
from storage.backend import get_storage_backend
from utils.cleanup import CleanupService
from utils.logging_config import configure_logging
from utils.slack import SlackNotifier
from worker.loop import WorkerLoop
from worker.pipeline import Pipeline

logger = logging.getLogger(__name__)


def build_pipeline(settings: Settings, backend: BackendClient) -> Pipeline:
    if sys.platform == "win32" and settings.ispring_adapter != "fake":
        powerpoint = PowerPointService(settings)
        browser = PlaywrightChecker(settings)
    else:
        powerpoint = NoOpPowerPointService()
        browser = (
            NoOpBrowserChecker()
            if settings.ispring_adapter == "fake" or not settings.browser_test_enabled
            else PlaywrightChecker(settings)
        )
        if sys.platform != "win32":
            browser = NoOpBrowserChecker()
    return Pipeline(
        settings,
        backend=backend,
        powerpoint=powerpoint,
        publisher=get_publisher(settings),
        storage=get_storage_backend(settings),
        browser=browser,
        cleanup=CleanupService(settings),
        slack=SlackNotifier(settings),
    )


def main() -> int:
    settings = get_settings()
    configure_logging(settings)
    if not settings.backend_base_url:
        logger.error("BACKEND_BASE_URL is required")
        return 2

    backend = BackendClient(settings)
    pipeline = build_pipeline(settings, backend)
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
            "storage_backend": settings.storage_backend,
            "ispring_adapter": settings.ispring_adapter,
        },
    )
    loop.run_forever()
    logger.info("worker stopped")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
