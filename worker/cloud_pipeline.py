"""One claimed PHP job: download, publish to iSpring Cloud, POST /result.

Never calls /retry. After success or failure the poll loop GETs /next again.
"""

from __future__ import annotations

import logging
from collections.abc import Callable

from api_client.jobs import Job, error_code_for, result_stage
from config.settings import Settings
from utils.exceptions import CallbackError, PptAutomationError
from utils.logging_config import bind_job_context, clear_job_context, set_stage
from worker import health
from worker.cloud_job import CloudJobResult, run_cloud_job

logger = logging.getLogger(__name__)

CloudJobFn = Callable[..., CloudJobResult]


class CloudPipeline:
    def __init__(
        self,
        settings: Settings,
        backend,
        *,
        cloud_job: CloudJobFn = run_cloud_job,
        slack=None,
    ) -> None:
        self._settings = settings
        self._backend = backend
        self._cloud_job = cloud_job
        self._slack = slack

    def process(self, job: Job) -> None:
        bind_job_context(job_id=job.job_id, material_id=job.material_id)
        error: BaseException | None = None
        result: CloudJobResult | None = None
        try:
            set_stage("download")
            result = self._cloud_job(
                url=job.file_url,
                material_name=job.material_name,
                material_id=job.material_id if job.material_id is not None else job.job_id,
                institution_name=job.institution_name or "",
                settings=self._settings,
                job_id=job.job_id,
                content_name="",
            )
            set_stage("callback")
            self._backend.send_job_result(
                job,
                status=2,
                ispringcloud_link=result.iframe_url,
            )
            health.record_success(self._settings.log_root)
        except PptAutomationError as exc:
            error = exc
            logger.error("pipeline failed", extra={"stage": exc.stage}, exc_info=True)
        except Exception as exc:
            error = PptAutomationError(str(exc), stage="unexpected")
            logger.error(
                "pipeline failed unexpectedly",
                extra={"stage": "unexpected"},
                exc_info=True,
            )
        finally:
            if error is not None:
                self._report_failure(job, error)
            clear_job_context()

    def _report_failure(self, job: Job, error: BaseException) -> None:
        stage = result_stage(getattr(error, "stage", "unexpected"))
        code = error_code_for(error)
        message = getattr(error, "user_message", None) or str(error)
        try:
            self._backend.send_job_result(
                job,
                status=3,
                stage=stage,
                error_code=code,
                error_message=message[:4000],
            )
        except CallbackError:
            logger.error(
                "failure callback itself failed",
                extra={"stage": "callback", "job_id": job.job_id},
                exc_info=True,
            )
        health.record_failure(
            self._settings.log_root,
            stage=getattr(error, "stage", "") or "",
            message=getattr(error, "message", None) or str(error),
        )
        if self._slack is not None:
            self._slack.notify_failure(job, stage=stage, error_message=message)
