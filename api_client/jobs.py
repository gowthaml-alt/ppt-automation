"""PHP job client for the iSpring Cloud worker.

Talks only to:

- GET  {base}/pptupdate/ispringcloud/next
- POST {base}/pptupdate/ispringcloud/result

Never calls POST /pptupdate/ispringcloud/retry. That is an ops API.
"""

from __future__ import annotations

import json
import logging
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urljoin

import httpx

from config.settings import Settings
from utils.exceptions import CallbackError, JobFetchError
from utils.http import create_sync_client
from utils.logging_config import scrub_url

logger = logging.getLogger(__name__)

NEXT_PATH = "/pptupdate/ispringcloud/next"
RESULT_PATH = "/pptupdate/ispringcloud/result"
RETRY_PATH = "/pptupdate/ispringcloud/retry"
TOKEN_HEADER = "X-PPT-WORKER-TOKEN"
# Password for every PHP queue call. Not read from env. Never invent another.
WORKER_TOKEN = "EdmPptWk_93e104bb70c17cc560f21e0b46e80bdb"
SUCCESS_CODE = "success"
# Resource::createJson() puts 200 in "code", not a word.
HTTP_OK_CODE = 200

RESULT_STAGES = frozenset({"download", "ispring", "upload", "callback"})

_INTERNAL_TO_RESULT_STAGE = {
    "fetch": "download",
    "prepare": "download",
    "download": "download",
    "powerpoint": "ispring",
    "powerpoint_open": "ispring",
    "ispring_publish": "ispring",
    "ispring": "ispring",
    "powerpoint_close": "ispring",
    "output_validate": "upload",
    "upload": "upload",
    "browser_test": "upload",
    "callback": "callback",
    "cleanup": "callback",
    "unexpected": "ispring",
}

_ERROR_CODES = {
    "DownloadError": "DOWNLOAD_FAILED",
    "LowDiskSpaceError": "DOWNLOAD_FAILED",
    "JobFetchError": "DOWNLOAD_FAILED",
    "PowerPointAutomationError": "ISPRING_FAILED",
    "ISpringPublishingError": "ISPRING_FAILED",
    "ISpringNotConfiguredError": "ISPRING_FAILED",
    "ISpringTimeoutError": "ISPRING_FAILED",
    "OutputValidationError": "UPLOAD_FAILED",
    "OutputUploadError": "UPLOAD_FAILED",
    "BrowserTestError": "UPLOAD_FAILED",
    "CallbackError": "CALLBACK_FAILED",
}


@dataclass(frozen=True)
class Job:
    job_id: int
    file_url: str
    asset_id: int | None = None
    material_id: int | None = None
    material_name: str = ""
    institution_id: int | None = None
    institution_name: str | None = None

    @property
    def queue_id(self) -> int:
        return self.job_id

    @property
    def ppt_file_url(self) -> str:
        return self.file_url


@dataclass(frozen=True)
class UploadResult:
    output_path: str
    entry_file: str
    iframe_url: str


def _join_url(base: str, path: str) -> str:
    return urljoin(base.rstrip("/") + "/", path.lstrip("/"))


def _is_success_code(code: Any) -> bool:
    """Did the API say this call worked?

    PHP's Resource::createJson() answers {"code": 200, "message": "Success"},
    so the integer 200 is the shape that actually arrives. The string form is
    still accepted for anything that reports "success" instead.
    """
    if isinstance(code, bool):
        return False
    if isinstance(code, int):
        return code == HTTP_OK_CODE
    if isinstance(code, str):
        stripped = code.strip()
        return stripped.lower() == SUCCESS_CODE or stripped == str(HTTP_OK_CODE)
    return False


def _optional_int(value: Any) -> int | None:
    if value is None or value == "":
        return None
    return int(value)


def result_stage(stage: str | None) -> str:
    if not stage:
        return "ispring"
    if stage in RESULT_STAGES:
        return stage
    return _INTERNAL_TO_RESULT_STAGE.get(stage, "ispring")


def error_code_for(error: BaseException) -> str:
    return _ERROR_CODES.get(type(error).__name__, "UNEXPECTED")


def parse_next_response(payload: Any) -> Job | None:
    """Map GET /next JSON onto :class:`Job`. Extra keys are ignored."""
    if not isinstance(payload, dict):
        raise JobFetchError(
            f"GET next payload must be an object, got {type(payload).__name__}",
            user_message="The backend returned an unexpected job payload.",
        )
    if not _is_success_code(payload.get("code")):
        raise JobFetchError(
            f"GET next returned code={payload.get('code')!r}",
            user_message="The backend job API did not return success.",
        )
    # createJson() merges its nodes into the top level, so "material" arrives
    # beside "code". A "data" wrapper is still read when one is present.
    if "material" in payload:
        material = payload.get("material")
    else:
        data = payload.get("data")
        if data is None:
            return None
        if not isinstance(data, dict):
            raise JobFetchError(
                f"GET next data must be an object, got {type(data).__name__}",
                user_message="The backend returned an unexpected job payload.",
            )
        material = data.get("material")
    if material is None:
        return None
    if not isinstance(material, dict):
        raise JobFetchError(
            f"GET next material must be an object, got {type(material).__name__}",
            user_message="The backend returned an unexpected job payload.",
        )
    return _parse_material(material)


def _parse_material(data: dict[str, Any]) -> Job:
    job_id = data.get("job_id")
    file_url = data.get("file_url")
    if job_id is None:
        raise JobFetchError(
            "GET next material missing job_id",
            user_message="The backend job payload is missing required fields.",
        )
    if not file_url:
        raise JobFetchError(
            "GET next material missing file_url",
            user_message="The backend job payload is missing required fields.",
        )
    try:
        job_id_int = int(job_id)
    except (TypeError, ValueError) as exc:
        raise JobFetchError(
            f"GET next material has a non-integer job_id: {exc}",
            user_message="The backend job payload has invalid ids.",
        ) from exc
    try:
        asset_id = _optional_int(data.get("asset_id"))
        material_id = _optional_int(data.get("material_id"))
        institution_id = _optional_int(data.get("institution_id"))
    except (TypeError, ValueError) as exc:
        raise JobFetchError(
            f"GET next material has a non-integer id: {exc}",
            user_message="The backend job payload has invalid ids.",
        ) from exc
    institution = data.get("institution_name")
    material_name = data.get("material_name") or ""
    return Job(
        job_id=job_id_int,
        file_url=str(file_url),
        asset_id=asset_id,
        material_id=material_id,
        material_name=str(material_name),
        institution_id=institution_id,
        institution_name=None if institution is None else str(institution),
    )


def record_orphaned_output(
    log_root: str | Path, job: Job, result: UploadResult
) -> Path:
    """Persist a successful publish that could not be reported to PHP."""
    destination = Path(log_root) / "orphaned_outputs.jsonl"
    destination.parent.mkdir(parents=True, exist_ok=True)
    row = {
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "job_id": job.job_id,
        "material_id": job.material_id,
        "iframe_url": result.iframe_url,
        "output_path": result.output_path,
        "entry_file": result.entry_file,
    }
    with destination.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(row) + "\n")
    return destination


class BackendClient:
    def __init__(
        self, settings: Settings, *, client: httpx.Client | None = None
    ) -> None:
        self._settings = settings
        self._client = client

    def _headers(self) -> dict[str, str]:
        return {
            "Accept": "application/json",
            TOKEN_HEADER: WORKER_TOKEN,
        }

    def _client_or_new(self) -> tuple[httpx.Client, bool]:
        if self._client is not None:
            return self._client, False
        return (
            create_sync_client(timeout_s=self._settings.backend_request_timeout_seconds),
            True,
        )

    def get_next_job(self) -> Job | None:
        if not self._settings.backend_base_url:
            raise JobFetchError(
                "BACKEND_BASE_URL is not configured",
                user_message="The backend URL is not configured.",
            )
        url = _join_url(self._settings.backend_base_url, NEXT_PATH)
        logger.info("GET next job", extra={"url": scrub_url(url), "stage": "fetch"})
        attempts = max(1, self._settings.get_job_retry_attempts)
        last_error: Exception | None = None
        client, owns = self._client_or_new()
        try:
            for attempt in range(1, attempts + 1):
                try:
                    response = client.get(
                        url,
                        headers=self._headers(),
                        params={"token": WORKER_TOKEN},
                    )
                except httpx.HTTPError as exc:
                    last_error = exc
                    logger.warning(
                        "GET next job transport failed",
                        extra={
                            "url": scrub_url(url),
                            "stage": "fetch",
                            "attempt": attempt,
                        },
                        exc_info=True,
                    )
                    if attempt < attempts:
                        time.sleep(min(2 ** (attempt - 1), 8))
                        continue
                    raise JobFetchError(
                        f"GET next job failed: {exc}",
                        user_message="The backend job API could not be reached.",
                    ) from exc
                if response.status_code != 200:
                    last_error = JobFetchError(
                        f"GET next job returned HTTP {response.status_code}",
                        user_message="The backend job API failed.",
                    )
                    logger.warning(
                        "GET next job HTTP error",
                        extra={
                            "url": scrub_url(url),
                            "stage": "fetch",
                            "status": response.status_code,
                            "attempt": attempt,
                        },
                    )
                    if attempt < attempts and response.status_code >= 500:
                        time.sleep(min(2 ** (attempt - 1), 8))
                        continue
                    raise last_error
                if not response.content:
                    raise JobFetchError(
                        "GET next job returned an empty body",
                        user_message="The backend job API returned invalid JSON.",
                    )
                try:
                    payload = response.json()
                except ValueError as exc:
                    raise JobFetchError(
                        f"GET next job returned non-JSON: {exc}",
                        user_message="The backend job API returned invalid JSON.",
                    ) from exc
                job = parse_next_response(payload)
                logger.info(
                    "got a job" if job is not None else "queue is empty",
                    extra={
                        "stage": "fetch",
                        "status": response.status_code,
                        "job_id": job.job_id if job is not None else None,
                    },
                )
                return job
        finally:
            if owns:
                client.close()
        if last_error:
            raise last_error
        return None

    def send_job_result(
        self,
        job: Job,
        *,
        status: int,
        ispringcloud_link: str | None = None,
        stage: str | None = None,
        error_code: str | None = None,
        error_message: str | None = None,
    ) -> None:
        if not self._settings.backend_base_url:
            raise CallbackError(
                "BACKEND_BASE_URL is not configured",
                user_message="The backend URL is not configured.",
            )
        if status not in (2, 3):
            raise CallbackError(
                f"result status must be 2 or 3, got {status!r}",
                user_message="The worker produced an invalid result status.",
            )
        url = _join_url(self._settings.backend_base_url, RESULT_PATH)
        logger.info(
            "POST job result",
            extra={"url": scrub_url(url), "stage": "callback", "status": status},
        )
        if status == 2:
            payload: dict[str, Any] = {
                "token": WORKER_TOKEN,
                "job_id": job.job_id,
                "status": 2,
                "ispringcloud_link": ispringcloud_link or "",
            }
        else:
            payload = {
                "token": WORKER_TOKEN,
                "job_id": job.job_id,
                "status": 3,
                "stage": result_stage(stage),
                "error_code": error_code or "UNEXPECTED",
                "error_message": error_message or "unknown error",
            }
        form = {"JSONString": json.dumps(payload)}

        headers = self._headers()
        attempts = max(1, self._settings.callback_retry_attempts)
        client, owns = self._client_or_new()
        last_error: Exception | None = None
        try:
            for attempt in range(1, attempts + 1):
                try:
                    response = client.post(url, data=form, headers=headers)
                except httpx.HTTPError as exc:
                    last_error = exc
                    logger.error(
                        "callback transport failed",
                        extra={
                            "url": scrub_url(url),
                            "stage": "callback",
                            "attempt": attempt,
                        },
                        exc_info=True,
                    )
                    if attempt < attempts:
                        time.sleep(min(2 ** (attempt - 1), 8))
                        continue
                    raise CallbackError(
                        f"callback failed: {exc}",
                        user_message="The backend callback API could not be reached.",
                    ) from exc
                if response.status_code >= 500:
                    last_error = CallbackError(
                        f"callback returned HTTP {response.status_code}",
                        user_message="The backend callback API failed.",
                    )
                    logger.error(
                        "callback server error",
                        extra={
                            "url": scrub_url(url),
                            "stage": "callback",
                            "status": response.status_code,
                            "attempt": attempt,
                        },
                    )
                    if attempt < attempts:
                        time.sleep(min(2 ** (attempt - 1), 8))
                        continue
                    raise last_error
                if response.status_code >= 400:
                    raise CallbackError(
                        f"callback returned HTTP {response.status_code}",
                        user_message="The backend callback API rejected the request.",
                    )
                logger.info(
                    "callback sent",
                    extra={"status": status, "stage": "callback", "job_id": job.job_id},
                )
                return
        finally:
            if owns:
                client.close()
        if last_error:
            raise last_error
