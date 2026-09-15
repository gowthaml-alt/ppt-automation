"""Internal job model and Node.js HTTP client.

This module talks only to HTTP APIs. It does not know the queue table schema
and never connects to MySQL.
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

_EMPTY_STATUSES = {204, 404}


@dataclass(frozen=True)
class Job:
    queue_id: int
    material_id: int
    ppt_file_url: str
    material_name: str = ""
    institution_name: str | None = None


@dataclass(frozen=True)
class UploadResult:
    output_path: str
    entry_file: str
    iframe_url: str


def _join_url(base: str, path: str) -> str:
    return urljoin(base.rstrip("/") + "/", path.lstrip("/"))


def _first(data: dict[str, Any], *keys: str) -> Any:
    for key in keys:
        if key in data and data[key] is not None and data[key] != "":
            return data[key]
    return None


def parse_job_payload(payload: Any) -> Job | None:
    """Map a GET body onto :class:`Job`. Extra keys are ignored."""
    data = payload
    if data is None:
        return None
    if isinstance(data, list):
        if not data:
            return None
        data = data[0]
    if not isinstance(data, dict):
        raise JobFetchError(
            f"GET job payload must be an object, got {type(data).__name__}",
            user_message="The backend returned an unexpected job payload.",
        )
    if not data:
        return None
    if "job" in data:
        nested = data["job"]
        if nested is None:
            return None
        if isinstance(nested, dict):
            data = nested
    elif "data" in data and not any(
        key in data for key in ("queue_id", "id", "job_id", "ppt_file_url", "material_id")
    ):
        nested = data["data"]
        if nested is None:
            return None
        if isinstance(nested, dict):
            data = nested
        elif isinstance(nested, list):
            if not nested:
                return None
            data = nested[0]

    queue_id = _first(data, "queue_id", "id", "job_id")
    material_id = _first(data, "material_id")
    ppt_file_url = _first(data, "ppt_file_url", "pptx_url", "file_url", "ppt_url")

    if queue_id is None and material_id is None and ppt_file_url is None:
        return None

    missing = []
    if queue_id is None:
        missing.append("queue_id")
    if material_id is None:
        missing.append("material_id")
    if not ppt_file_url:
        missing.append("ppt_file_url")
    if missing:
        raise JobFetchError(
            f"GET job payload missing required fields: {', '.join(missing)}",
            user_message="The backend job payload is missing required fields.",
        )

    try:
        queue_id_int = int(queue_id)
        material_id_int = int(material_id)
    except (TypeError, ValueError) as exc:
        raise JobFetchError(
            f"GET job payload has non-integer ids: {exc}",
            user_message="The backend job payload has invalid ids.",
        ) from exc

    institution = _first(data, "institution_name")
    material_name = _first(data, "material_name") or ""
    return Job(
        queue_id=queue_id_int,
        material_id=material_id_int,
        ppt_file_url=str(ppt_file_url),
        material_name=str(material_name),
        institution_name=None if institution is None else str(institution),
    )


def record_orphaned_output(
    log_root: str | Path, job: Job, result: UploadResult
) -> Path:
    """Persist a successful upload that could not be reported to Node.js."""
    destination = Path(log_root) / "orphaned_outputs.jsonl"
    destination.parent.mkdir(parents=True, exist_ok=True)
    row = {
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "queue_id": job.queue_id,
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
        headers = {"Accept": "application/json"}
        if self._settings.backend_api_key:
            headers[self._settings.backend_api_key_header] = (
                self._settings.backend_api_key
            )
        return headers

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
        url = _join_url(
            self._settings.backend_base_url, self._settings.backend_get_job_path
        )
        attempts = max(1, self._settings.get_job_retry_attempts)
        last_error: Exception | None = None
        client, owns = self._client_or_new()
        try:
            for attempt in range(1, attempts + 1):
                try:
                    response = client.get(url, headers=self._headers())
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
                if response.status_code in _EMPTY_STATUSES:
                    return None
                if response.status_code >= 500:
                    last_error = JobFetchError(
                        f"GET next job returned HTTP {response.status_code}",
                        user_message="The backend job API failed.",
                    )
                    logger.warning(
                        "GET next job server error",
                        extra={
                            "url": scrub_url(url),
                            "stage": "fetch",
                            "status": response.status_code,
                            "attempt": attempt,
                        },
                    )
                    if attempt < attempts:
                        time.sleep(min(2 ** (attempt - 1), 8))
                        continue
                    raise last_error
                if response.status_code >= 400:
                    raise JobFetchError(
                        f"GET next job returned HTTP {response.status_code}",
                        user_message="The backend job API rejected the request.",
                    )
                if not response.content:
                    return None
                try:
                    payload = response.json()
                except ValueError as exc:
                    raise JobFetchError(
                        f"GET next job returned non-JSON: {exc}",
                        user_message="The backend job API returned invalid JSON.",
                    ) from exc
                return parse_job_payload(payload)
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
        status: str,
        iframe_url: str | None = None,
        error_message: str | None = None,
    ) -> None:
        if not self._settings.backend_base_url:
            raise CallbackError(
                "BACKEND_BASE_URL is not configured",
                user_message="The backend URL is not configured.",
            )
        url = _join_url(
            self._settings.backend_base_url, self._settings.backend_callback_path
        )
        body: dict[str, Any] = {
            "queue_id": job.queue_id,
            "material_id": job.material_id,
            "status": status,
        }
        if status == "completed":
            body["iframe_url"] = iframe_url
        else:
            body["error_message"] = error_message or "unknown error"

        headers = {**self._headers(), "Content-Type": "application/json"}
        attempts = max(1, self._settings.callback_retry_attempts)
        client, owns = self._client_or_new()
        last_error: Exception | None = None
        try:
            for attempt in range(1, attempts + 1):
                try:
                    response = client.post(url, json=body, headers=headers)
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
                    extra={"status": status, "stage": "callback"},
                )
                return
        finally:
            if owns:
                client.close()
        if last_error:
            raise last_error
