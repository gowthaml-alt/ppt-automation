"""Outbound poll loop. Fetches one job at a time; never claims a database row."""

from __future__ import annotations

import logging
import time
from collections.abc import Callable
from typing import Protocol

from api_client.jobs import Job
from config.settings import Settings
from utils.exceptions import JobFetchError, PptAutomationError
from worker import health

logger = logging.getLogger(__name__)


class JobProcessor(Protocol):
    def process(self, job: Job) -> None: ...


class NextJobSource(Protocol):
    def get_next_job(self) -> Job | None: ...


class WorkerLoop:
    def __init__(
        self,
        settings: Settings,
        backend: NextJobSource,
        pipeline: JobProcessor,
        *,
        sleeper: Callable[[float], None] = time.sleep,
    ) -> None:
        self._settings = settings
        self._backend = backend
        self._pipeline = pipeline
        self._sleeper = sleeper
        self._busy = False
        self._stop = False

    @property
    def busy(self) -> bool:
        return self._busy

    def stop(self) -> None:
        self._stop = True

    def run_forever(self) -> None:
        logger.info("poll loop starting", extra={"stage": "fetch"})
        while not self._stop:
            self.run_once()

    def run_once(self) -> None:
        if self._busy:
            logger.warning(
                "run_once called while a job is already in flight; ignoring",
                extra={"stage": "fetch"},
            )
            return
        if not health.should_take_work(self._settings.log_root):
            self._sleeper(self._settings.poll_interval_seconds)
            return
        try:
            job = self._backend.get_next_job()
        except JobFetchError:
            logger.error("could not fetch next job", extra={"stage": "fetch"}, exc_info=True)
            self._sleeper(self._settings.poll_interval_seconds)
            return
        if job is None:
            logger.info("no job available", extra={"stage": "fetch"})
            self._sleeper(self._settings.poll_interval_seconds)
            return

        self._busy = True
        try:
            logger.info(
                "processing job",
                extra={
                    "job_id": job.job_id,
                    "material_id": job.material_id,
                    "stage": "fetch",
                },
            )
            self._pipeline.process(job)
        except PptAutomationError:
            # Failure callback already sent inside the pipeline.
            logger.error(
                "job ended in failure",
                extra={"job_id": job.job_id, "stage": "unexpected"},
                exc_info=True,
            )
        finally:
            self._busy = False
