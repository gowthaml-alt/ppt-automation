"""Stream an input PPT to the job directory. Never trusts the remote filename."""

from __future__ import annotations

import logging
from pathlib import Path

import httpx

from config.settings import Settings
from utils.exceptions import DownloadError
from utils.http import create_sync_client
from utils.logging_config import scrub_url
from utils.paths import JobPaths
from utils.validators import (
    assert_looks_like_pptx,
    assert_valid_pptx_package,
    validate_download_url,
)

logger = logging.getLogger(__name__)

CHUNK_SIZE = 64 * 1024


def download_pptx(
    url: str,
    paths: JobPaths,
    settings: Settings,
    *,
    client: httpx.Client | None = None,
) -> Path:
    validate_download_url(url, settings.download_allowed_hosts)
    dest = paths.source_pptx
    partial = paths.partial_pptx
    owns_client = client is None
    client = client or create_sync_client(timeout_s=settings.download_timeout_seconds)

    logger.info("download starting", extra={"url": scrub_url(url), "stage": "download"})
    try:
        _stream_to_partial(client, url, partial, settings.max_ppt_size_bytes)
        assert_looks_like_pptx(partial)
        assert_valid_pptx_package(partial)
        partial.replace(dest)
    except DownloadError:
        _remove_if_exists(partial)
        _remove_if_exists(dest)
        raise
    except httpx.TimeoutException as exc:
        _remove_if_exists(partial)
        _remove_if_exists(dest)
        raise DownloadError(
            f"download timed out for {scrub_url(url)}: {exc}",
            user_message="Downloading the PowerPoint file timed out.",
        ) from exc
    except httpx.HTTPError as exc:
        _remove_if_exists(partial)
        _remove_if_exists(dest)
        raise DownloadError(
            f"download HTTP error for {scrub_url(url)}: {exc}",
            user_message="The PowerPoint file could not be downloaded.",
        ) from exc
    finally:
        if owns_client:
            client.close()

    logger.info(
        "download finished",
        extra={"bytes": dest.stat().st_size, "stage": "download"},
    )
    return dest


def _stream_to_partial(
    client: httpx.Client, url: str, partial: Path, max_bytes: int
) -> None:
    with client.stream("GET", url) as response:
        if response.status_code >= 400:
            raise DownloadError(
                f"download returned HTTP {response.status_code} for {scrub_url(url)}",
                user_message="The storage server refused the PowerPoint download.",
            )
        written = 0
        with partial.open("wb") as handle:
            for chunk in response.iter_bytes(CHUNK_SIZE):
                written += len(chunk)
                if written > max_bytes:
                    raise DownloadError(
                        f"download exceeded max size {max_bytes} bytes "
                        f"({scrub_url(url)})",
                        user_message="The PowerPoint file is larger than the configured limit.",
                    )
                handle.write(chunk)


def _remove_if_exists(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except OSError:
        logger.warning("could not remove %s", path.name, extra={"stage": "download"}, exc_info=True)
