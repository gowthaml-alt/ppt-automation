"""Process one job through download, publish, upload, browser test, and callback."""

from __future__ import annotations

import logging
from collections.abc import Callable
from typing import Protocol

from api_client.jobs import Job, UploadResult, record_orphaned_output
from config.settings import Settings
from downloader.pptx import download_pptx
from storage.backend import output_prefix
from utils.cleanup import CleanupService
from utils.exceptions import CallbackError, LowDiskSpaceError, PptAutomationError
from utils.logging_config import bind_job_context, clear_job_context, set_stage
from utils.paths import JobPaths, build_job_paths, create_job_dirs, free_disk_gb
from validator.html5 import validate_ispring_output

logger = logging.getLogger(__name__)


class PowerPointLike(Protocol):
    def start(self) -> None: ...
    def open(self, path) -> None: ...
    def close(self) -> None: ...
    def quit(self) -> None: ...


class PublisherLike(Protocol):
    def publish(self, pptx, output_dir, timeout_s: int): ...


class StorageLike(Protocol):
    def upload_directory(self, local, prefix: str) -> UploadResult: ...


class BackendLike(Protocol):
    def send_job_result(
        self,
        job: Job,
        *,
        status: str,
        iframe_url: str | None = None,
        error_message: str | None = None,
    ) -> None: ...


class BrowserLike(Protocol):
    def check(self, iframe_url: str) -> None: ...


class SlackLike(Protocol):
    def notify_failure(self, job: Job, *, stage: str, error_message: str) -> None: ...


class Pipeline:
    def __init__(
        self,
        settings: Settings,
        *,
        backend: BackendLike,
        powerpoint: PowerPointLike,
        publisher: PublisherLike,
        storage: StorageLike,
        browser: BrowserLike,
        cleanup: CleanupService,
        slack: SlackLike | None = None,
        downloader: Callable = download_pptx,
    ) -> None:
        self._settings = settings
        self._backend = backend
        self._powerpoint = powerpoint
        self._publisher = publisher
        self._storage = storage
        self._browser = browser
        self._cleanup = cleanup
        self._slack = slack
        self._downloader = downloader

    def process(self, job: Job) -> None:
        bind_job_context(queue_id=job.queue_id, material_id=job.material_id)
        paths = build_job_paths(self._settings.temp_root, job.queue_id)
        original_error: BaseException | None = None
        upload_result: UploadResult | None = None
        success = False
        try:
            self._prepare(paths)
            set_stage("download")
            self._downloader(job.ppt_file_url, paths, self._settings)
            set_stage("powerpoint_open")
            self._powerpoint.start()
            self._powerpoint.open(paths.source_pptx)
            set_stage("ispring_publish")
            self._publisher.publish(
                paths.source_pptx,
                paths.output_dir,
                self._settings.ispring_publish_timeout_seconds,
            )
            set_stage("powerpoint_close")
            self._powerpoint.close()
            self._powerpoint.quit()
            set_stage("output_validate")
            validate_ispring_output(paths.output_dir)
            set_stage("upload")
            prefix = output_prefix(job.material_id, self._settings)
            upload_result = self._storage.upload_directory(paths.output_dir, prefix)
            set_stage("browser_test")
            self._browser.check(upload_result.iframe_url)
            set_stage("callback")
            self._backend.send_job_result(
                job, status="completed", iframe_url=upload_result.iframe_url
            )
            success = True
        except PptAutomationError as exc:
            original_error = exc
            logger.error(
                "pipeline failed",
                extra={"stage": exc.stage},
                exc_info=True,
            )
        except Exception as exc:
            original_error = PptAutomationError(str(exc), stage="unexpected")
            logger.error(
                "pipeline failed unexpectedly",
                extra={"stage": "unexpected"},
                exc_info=True,
            )
        finally:
            self._safe_close_powerpoint()
            if original_error is not None:
                self._report_failure(job, original_error, upload_result)
            self._cleanup.run(
                paths, success=success, original_error=original_error
            )
            clear_job_context()

        if original_error is not None:
            raise original_error

    def _prepare(self, paths: JobPaths) -> None:
        set_stage("prepare")
        free_gb = free_disk_gb(self._settings.temp_root)
        if free_gb < self._settings.min_free_disk_gb:
            raise LowDiskSpaceError(
                f"{free_gb:.1f} GiB free, {self._settings.min_free_disk_gb} GiB required",
                user_message="There is not enough disk space to process this job.",
            )
        create_job_dirs(paths)

    def _safe_close_powerpoint(self) -> None:
        try:
            self._powerpoint.close()
            self._powerpoint.quit()
        except Exception:
            logger.warning("PowerPoint teardown failed", exc_info=True)

    def _report_failure(
        self,
        job: Job,
        error: BaseException,
        upload_result: UploadResult | None,
    ) -> None:
        stage = getattr(error, "stage", "unexpected")
        message = getattr(error, "user_message", None) or str(error)
        if upload_result is not None:
            record_orphaned_output(self._settings.log_root, job, upload_result)
            message = f"{message} (uploaded iframe_url preserved in orphaned_outputs.jsonl)"
        try:
            self._backend.send_job_result(
                job, status="failed", error_message=message[:4000]
            )
        except CallbackError:
            logger.error(
                "failure callback itself failed",
                extra={"stage": "callback"},
                exc_info=True,
            )
        if self._slack is not None:
            self._slack.notify_failure(job, stage=str(stage), error_message=message)
