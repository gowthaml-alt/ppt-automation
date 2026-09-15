"""Storage factory, MIME map, and upload types."""

from __future__ import annotations

import logging
import mimetypes
import shutil
from pathlib import Path
from typing import Protocol

import httpx

from api_client.jobs import UploadResult
from config.settings import Settings
from utils.exceptions import OutputUploadError
from utils.http import create_sync_client
from utils.logging_config import scrub_url
from utils.paths import iter_files

logger = logging.getLogger(__name__)

MIME_BY_SUFFIX = {
    ".html": "text/html",
    ".htm": "text/html",
    ".js": "application/javascript",
    ".css": "text/css",
    ".json": "application/json",
    ".svg": "image/svg+xml",
    ".woff": "font/woff",
    ".woff2": "font/woff2",
    ".ttf": "font/ttf",
    ".mp4": "video/mp4",
    ".mp3": "audio/mpeg",
    ".wasm": "application/wasm",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".webp": "image/webp",
    ".ico": "image/x-icon",
}


class StorageBackend(Protocol):
    def upload_directory(self, local: Path, prefix: str) -> UploadResult: ...

    def verify(self, entry_url: str, *, client: httpx.Client | None = None) -> bool: ...


def mime_type_for(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in MIME_BY_SUFFIX:
        return MIME_BY_SUFFIX[suffix]
    guessed, _ = mimetypes.guess_type(path.name)
    return guessed or "application/octet-stream"


def output_prefix(material_id: int, settings: Settings) -> str:
    parts = [settings.s3_key_prefix.strip("/"), "ppt", str(material_id)]
    return "/".join(part for part in parts if part)


def head_or_range(
    entry_url: str, timeout_s: float, client: httpx.Client | None = None
) -> bool:
    owns = client is None
    client = client or create_sync_client(timeout_s=timeout_s)
    try:
        head = client.head(entry_url)
        if head.status_code < 400:
            return True
        ranged = client.get(entry_url, headers={"Range": "bytes=0-63"})
        return ranged.status_code < 400
    except httpx.HTTPError:
        logger.warning(
            "storage verify failed",
            extra={"url": scrub_url(entry_url), "stage": "upload"},
            exc_info=True,
        )
        return False
    finally:
        if owns:
            client.close()


class LocalFsStorageBackend:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings

    def upload_directory(self, local: Path, prefix: str) -> UploadResult:
        root = self._settings.local_storage_root
        if not root:
            raise OutputUploadError(
                "LOCAL_STORAGE_ROOT is not configured",
                user_message="Storage is not configured.",
            )
        destination = Path(root) / prefix
        if destination.exists():
            shutil.rmtree(destination)
        shutil.copytree(local, destination)
        base = self._settings.local_storage_public_base_url.rstrip("/")
        iframe_url = f"{base}/{prefix.strip('/')}/index.html"
        logger.info(
            "local_fs upload finished", extra={"stage": "upload", "prefix": prefix}
        )
        return UploadResult(
            output_path=prefix.strip("/"),
            entry_file="index.html",
            iframe_url=iframe_url,
        )

    def verify(self, entry_url: str, *, client: httpx.Client | None = None) -> bool:
        return head_or_range(entry_url, self._settings.upload_timeout_seconds, client)


class S3StorageBackend:
    def __init__(self, settings: Settings, client=None) -> None:
        self._settings = settings
        self._client = client

    def _s3(self):
        if self._client is not None:
            return self._client
        import boto3
        from botocore.config import Config

        kwargs = {
            "service_name": "s3",
            "region_name": self._settings.s3_region or None,
            "aws_access_key_id": self._settings.s3_access_key_id or None,
            "aws_secret_access_key": self._settings.s3_secret_access_key or None,
            "config": Config(retries={"max_attempts": 3, "mode": "standard"}),
        }
        if self._settings.s3_endpoint_url:
            kwargs["endpoint_url"] = self._settings.s3_endpoint_url
        return boto3.client(**kwargs)

    def upload_directory(self, local: Path, prefix: str) -> UploadResult:
        if not self._settings.s3_bucket:
            raise OutputUploadError(
                "S3_BUCKET is not configured",
                user_message="Storage is not configured.",
            )
        if not self._settings.s3_public_base_url:
            raise OutputUploadError(
                "S3_PUBLIC_BASE_URL is not configured",
                user_message="Storage is not configured.",
            )
        client = self._s3()
        try:
            for absolute, relative in iter_files(local):
                key = f"{prefix}/{relative}"
                client.upload_file(
                    str(absolute),
                    self._settings.s3_bucket,
                    key,
                    ExtraArgs={"ContentType": mime_type_for(absolute)},
                )
        except Exception as exc:
            logger.error("s3 upload failed", extra={"stage": "upload"}, exc_info=True)
            raise OutputUploadError(
                f"s3 upload failed: {exc}",
                user_message="The published package could not be uploaded.",
            ) from exc
        iframe_url = (
            f"{self._settings.s3_public_base_url.rstrip('/')}/{prefix}/index.html"
        )
        logger.info(
            "s3 upload finished",
            extra={
                "stage": "upload",
                "prefix": prefix,
                "url": scrub_url(iframe_url),
            },
        )
        return UploadResult(
            output_path=prefix,
            entry_file="index.html",
            iframe_url=iframe_url,
        )

    def verify(self, entry_url: str, *, client: httpx.Client | None = None) -> bool:
        return head_or_range(entry_url, self._settings.upload_timeout_seconds, client)


class HttpApiStorageBackend:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings

    def upload_directory(self, local: Path, prefix: str) -> UploadResult:
        raise OutputUploadError(
            "STORAGE_BACKEND=http_api is not implemented. Needed facts: "
            "endpoint path, authentication scheme, whether the upload is multipart "
            "files or a zip, and the success response shape including the iframe URL.",
            user_message="HTTP API storage is not implemented.",
        )

    def verify(self, entry_url: str, *, client: httpx.Client | None = None) -> bool:
        raise OutputUploadError(
            "STORAGE_BACKEND=http_api is not implemented.",
            user_message="HTTP API storage is not implemented.",
        )


def get_storage_backend(settings: Settings) -> StorageBackend:
    if settings.storage_backend == "local_fs":
        return LocalFsStorageBackend(settings)
    if settings.storage_backend == "s3":
        return S3StorageBackend(settings)
    if settings.storage_backend == "http_api":
        return HttpApiStorageBackend(settings)
    raise OutputUploadError(
        f"unknown STORAGE_BACKEND {settings.storage_backend!r}",
        user_message="Storage is not configured.",
    )
