"""One job: a PPT URL in, an iSpring Cloud iframe URL out.

    result = run_cloud_job(
        url="https://cdn.example.com/...Kickoff.pptx?Expires=...",
        material_name="Kickoff and Advanced Prompting",
        material_id=20242897,
        institution_name="Demoacademy",
        settings=get_settings(),
    )
    result.iframe_url

The download is deliberately forgiving about damaged files. A deck whose zip
is broken or whose parts are missing is still handed to PowerPoint, because
PowerPoint's repair is the only thing that can fix it and it usually does.
Only two things fail before PowerPoint is involved: a file that is not a
presentation at all, and one that is password-protected — no amount of
repairing gets past a password.
"""

from __future__ import annotations

import logging
import shutil
import time
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import unquote, urlsplit

from config.settings import Settings
from publisher.ispring_cloud import CloudPublishResult, RepairPromptWatcher, publish_to_cloud
from utils.exceptions import DownloadError, PowerPointAutomationError
from utils.http import create_sync_client
from utils.logging_config import scrub_url
from utils.validators import (
    USER_PASSWORD_PROTECTED,
    assert_valid_pptx_package,
    validate_download_url,
)

logger = logging.getLogger(__name__)

CHUNK_SIZE = 64 * 1024
OOXML_MAGIC = b"PK\x03\x04"  # .pptx and friends
OLE_MAGIC = b"\xd0\xcf\x11\xe0"  # the old binary .ppt
SAFE_NAME = "source"


@dataclass(frozen=True)
class CloudJobResult:
    iframe_url: str
    embed_code: str
    content_name: str
    source_name: str
    repaired: bool
    elapsed_s: float


def filename_from_url(url: str) -> str:
    """The file name a signed URL points at, ignoring its query string.

    Signed CloudFront URLs carry the signature in the query, and the path can
    hold brackets and percent-escapes, so the name is taken from the path
    alone and only used for logging and the extension.
    """
    path = urlsplit(url or "").path
    name = unquote(path.rsplit("/", 1)[-1]) if path else ""
    return name.strip() or "download.pptx"


def _extension_for(name: str) -> str:
    suffix = Path(name).suffix.lower()
    return suffix if suffix in {".pptx", ".ppt", ".pptm", ".potx"} else ".pptx"


def download(url: str, into: Path, settings: Settings) -> Path:
    """Stream the deck to disk. Never trusts the remote file name."""
    validate_download_url(url, settings.download_allowed_hosts)
    source_name = filename_from_url(url)
    target = into / f"{SAFE_NAME}{_extension_for(source_name)}"
    partial = into / f"{SAFE_NAME}.part"
    client = create_sync_client(timeout_s=settings.download_timeout_seconds)
    logger.info(
        "download starting",
        extra={"stage": "download", "url": scrub_url(url), "source_name": source_name},
    )
    try:
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
                    if written > settings.max_ppt_size_bytes:
                        raise DownloadError(
                            f"download exceeded {settings.max_ppt_size_bytes} bytes",
                            user_message="The PowerPoint file is larger than the limit.",
                        )
                    handle.write(chunk)
    finally:
        client.close()

    if not partial.exists() or partial.stat().st_size == 0:
        raise DownloadError(
            "downloaded file is empty", user_message="The downloaded file is empty."
        )
    partial.replace(target)
    logger.info(
        "download finished",
        extra={"stage": "download", "bytes": target.stat().st_size},
    )
    return target


def inspect(pptx: Path) -> bool:
    """Decide whether this file is worth handing to PowerPoint.

    Returns True when the package looks damaged, so the caller knows to expect
    a repair. Raises only for files PowerPoint cannot rescue.
    """
    with pptx.open("rb") as handle:
        header = handle.read(4)
    if header not in (OOXML_MAGIC, OLE_MAGIC):
        raise DownloadError(
            f"downloaded file is not a presentation (header {header!r})",
            user_message="The downloaded file is not a PowerPoint file.",
        )
    if header == OLE_MAGIC:
        logger.info(
            "old binary .ppt format; PowerPoint will convert it",
            extra={"stage": "download"},
        )
        return False

    try:
        assert_valid_pptx_package(pptx)
    except DownloadError as exc:
        if exc.user_message == USER_PASSWORD_PROTECTED:
            raise  # a password is not something repair can undo
        logger.warning(
            "package looks damaged; will rely on PowerPoint repair",
            extra={"stage": "download", "detail": str(exc)},
        )
        return True
    return False


def open_with_repair(pptx: Path, settings: Settings):
    """Open the deck through PowerPoint, repairing it if need be.

    Uses the service's OpenAndRepair path, with a watcher answering the repair
    prompt from another thread — PowerPoint's COM open does not return while
    that dialog is up.

    Returns ``(service, path_in_use)``. A repaired deck is saved under a new
    name, so the path in use is not always the path passed in.
    """
    from powerpoint.service import PowerPointService

    service = PowerPointService(settings)
    service.start()
    with RepairPromptWatcher() as watcher:
        service.open(pptx, timeout_s=settings.powerpoint_open_timeout_seconds)
    if watcher.answered:
        logger.info(
            "PowerPoint repaired the deck",
            extra={"stage": "powerpoint_open", "prompts_answered": watcher.answered},
        )
    in_use = service.current_pptx or pptx
    return service, Path(in_use)


def remove_workspace(workspace: Path, attempts: int = 5) -> bool:
    """Delete the downloaded and repaired files.

    Windows keeps the file locked for a moment after PowerPoint quits, so the
    first delete can fail even though nothing is really using the file.
    """
    for attempt in range(1, attempts + 1):
        shutil.rmtree(workspace, ignore_errors=True)
        if not workspace.exists():
            logger.info(
                "job files deleted",
                extra={"stage": "cleanup", "folder": workspace.name},
            )
            return True
        time.sleep(1)
    logger.warning(
        "could not delete the job files; something still has them open",
        extra={"stage": "cleanup", "folder": str(workspace)},
    )
    return False


def run_cloud_job(
    *,
    url: str = "",
    pptx: str | Path = "",
    material_name: str,
    material_id: int | str,
    institution_name: str,
    settings: Settings,
    content_name: str = "",
    keep_files: bool = False,
) -> CloudJobResult:
    """Download, repair if needed, publish, and return the iframe URL."""
    if not url and not pptx:
        raise ValueError("pass either url or pptx")
    started = time.monotonic()
    workspace = Path(settings.temp_root) / f"cloud-{material_id}"
    if workspace.exists():
        shutil.rmtree(workspace, ignore_errors=True)
    workspace.mkdir(parents=True, exist_ok=True)

    service = None
    published = False
    try:
        if url:
            source_name = filename_from_url(url)
            deck = download(url, workspace, settings)
        else:
            original = Path(pptx).expanduser().resolve()
            source_name = original.name
            deck = workspace / f"{SAFE_NAME}{_extension_for(original.name)}"
            shutil.copy2(original, deck)

        damaged = inspect(deck)
        service, in_use = open_with_repair(deck, settings)
        repaired = damaged or in_use != deck

        result: CloudPublishResult = publish_to_cloud(
            in_use,
            institution=institution_name,
            content_name=content_name or material_name,
            parent_folder=settings.ispring_parent_folder,
            cdp_url=settings.ispring_chrome_cdp_url,
            publish_timeout_s=settings.ispring_publish_timeout_seconds,
            close_powerpoint_after=False,
            skip_open=True,
        )
        published = True
    finally:
        # PowerPoint has to let go of the file before it can be deleted.
        if service is not None:
            try:
                service.quit()
            except Exception:  # noqa: BLE001
                logger.warning("PowerPoint teardown failed", exc_info=True)
        if keep_files:
            logger.info(
                "keeping the job files (--keep-files)",
                extra={"stage": "cleanup", "folder": str(workspace)},
            )
        elif published or not settings.keep_failed_job_files:
            remove_workspace(workspace)
        else:
            # A failed job's files are worth keeping: they are the evidence.
            logger.info(
                "keeping the files of a failed job (KEEP_FAILED_JOB_FILES)",
                extra={"stage": "cleanup", "folder": str(workspace)},
            )

    return CloudJobResult(
        iframe_url=result.iframe_url,
        embed_code=result.embed_code,
        content_name=content_name or material_name,
        source_name=source_name,
        repaired=repaired,
        elapsed_s=time.monotonic() - started,
    )
