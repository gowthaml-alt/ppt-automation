"""Delete job-scoped temporary files. Always invoked from a finally block."""

from __future__ import annotations

import logging
import shutil
from pathlib import Path

from config.settings import Settings
from utils.exceptions import CleanupError
from utils.paths import JobPaths

logger = logging.getLogger(__name__)


class CleanupService:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings

    def run(
        self,
        paths: JobPaths,
        *,
        success: bool,
        original_error: BaseException | None = None,
    ) -> None:
        """Best-effort cleanup. Never raises."""
        try:
            self._run(paths, success=success)
        except Exception as exc:
            logger.error(
                "cleanup failed",
                extra={"stage": "cleanup", "success": success},
                exc_info=True,
            )
            if original_error is None:
                logger.error(
                    "cleanup failed with no prior processing error: %s",
                    CleanupError(
                        str(exc),
                        user_message="Temporary files could not be deleted.",
                    ),
                    exc_info=True,
                )

    def _run(self, paths: JobPaths, *, success: bool) -> None:
        if not paths.root.exists():
            return
        if success:
            shutil.rmtree(paths.root)
            logger.info("job directory deleted", extra={"stage": "cleanup"})
            return
        self._preserve_logs(paths)
        if self._settings.keep_failed_job_files:
            logger.info(
                "KEEP_FAILED_JOB_FILES is set; leaving the job directory",
                extra={"stage": "cleanup"},
            )
            return
        shutil.rmtree(paths.root)
        logger.info(
            "failed-job directory deleted after log copy", extra={"stage": "cleanup"}
        )

    def _preserve_logs(self, paths: JobPaths) -> None:
        if not paths.log_dir.exists():
            return
        destination = Path(self._settings.log_root) / "jobs" / paths.root.name
        destination.mkdir(parents=True, exist_ok=True)
        for item in paths.log_dir.iterdir():
            if item.is_file():
                shutil.copy2(item, destination / item.name)
