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

# One private server cannot hold the loop for ever. Whatever is left over
# waits for the next round, ten minutes later.
PRIVATE_MAX_PER_ROUND = 25


class JobProcessor(Protocol):
    def process(self, job: Job) -> None: ...


class NextJobSource(Protocol):
    def get_next_job(self, base_url: str | None = None) -> Job | None: ...


class WorkerLoop:
    def __init__(
        self,
        settings: Settings,
        backend: NextJobSource,
        pipeline: JobProcessor,
        *,
        sleeper: Callable[[float], None] = time.sleep,
        clock: Callable[[], float] = time.monotonic,
    ) -> None:
        self._settings = settings
        self._backend = backend
        self._pipeline = pipeline
        self._sleeper = sleeper
        self._clock = clock
        self._busy = False
        self._stop = False
        # Private-server institutions keep their queue on their own host.
        # Due straight away, then once every private_poll_interval_seconds.
        self._private_urls = list(settings.private_base_urls)
        self._private_due = clock()

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
        # The private servers get their turn on their own clock. One round,
        # then back here, so the shared queue is never left waiting on them.
        if self._private_urls and self._clock() >= self._private_due:
            self._poll_private()
            return
        job = self._fetch(None)
        if job is None:
            self._sleeper(self._settings.poll_interval_seconds)
            return
        self._run(job)

    def _poll_private(self) -> None:
        """One round of the private servers: empty each one, then the next."""
        for base_url in self._private_urls:
            taken = 0
            while not self._stop and taken < PRIVATE_MAX_PER_ROUND:
                if not health.should_take_work(self._settings.log_root):
                    break
                job = self._fetch(base_url)
                if job is None:
                    break
                self._run(job)
                taken += 1
            if taken:
                logger.info(
                    "private server done for this round",
                    extra={"stage": "fetch", "server": base_url, "jobs": taken},
                )
            if self._stop:
                break
        self._private_due = (
            self._clock() + self._settings.private_poll_interval_seconds
        )

    def _fetch(self, base_url: str | None) -> Job | None:
        """One GET /next. A server that is down is logged, never raised."""
        try:
            if base_url is None:
                job = self._backend.get_next_job()
            else:
                job = self._backend.get_next_job(base_url)
        except JobFetchError:
            logger.error(
                "could not fetch next job",
                extra={"stage": "fetch", "server": base_url or "shared"},
                exc_info=True,
            )
            return None
        if job is None:
            logger.info(
                "no job available",
                extra={"stage": "fetch", "server": base_url or "shared"},
            )
        return job

    def _run(self, job: Job) -> None:
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
