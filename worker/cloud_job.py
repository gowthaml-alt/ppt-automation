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
from publisher.ispring_cloud import (
    CloudPublishResult,
    RepairPromptWatcher,
    ensure_project_folder,
    publish_to_cloud,
)
from utils.exceptions import (
    DownloadError,
    PowerPointAutomationError,
    ProjectMissingError,
)
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


def powerpoint_is_running() -> bool:
    try:
        import psutil  # type: ignore
    except ImportError:
        return True  # cannot tell; assume it is
    for process in psutil.process_iter(["name"]):
        try:
            if (process.info.get("name") or "").lower() == "powerpnt.exe":
                return True
        except Exception:  # noqa: BLE001
            continue
    return False


def powerpoint_answers_com() -> bool:
    """Is the running PowerPoint still usable through COM?

    A half-dead instance hands back an Application object whose Presentations
    collection has no Open2007, and every later call fails oddly. That is the
    thing worth checking before deciding to kill it.
    """
    try:
        import win32com.client  # type: ignore

        app = win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception:  # noqa: BLE001
        return False
    try:
        return getattr(getattr(app, "Presentations", None), "Open2007", None) is not None
    except Exception:  # noqa: BLE001
        return False


def ensure_usable_powerpoint(terminate) -> None:
    """Only kill PowerPoint when it has actually stopped working.

    Killing it is not free: Office treats a killed process as a crash and moves
    whatever add-in was loaded into Disabled Items, which is how the iSpring
    ribbon tab disappears. So a healthy instance is left alone.
    """
    from publisher.ispring_cloud import ensure_com

    ensure_com()
    if not powerpoint_is_running():
        return
    if powerpoint_answers_com():
        logger.info(
            "reusing the PowerPoint that is already running",
            extra={"stage": "powerpoint_open"},
        )
        return
    logger.warning(
        "the running PowerPoint is not answering COM; closing it. "
        "Office may disable the iSpring add-in after this — "
        "scripts/fix_ispring_addin.py --fix puts it back",
        extra={"stage": "powerpoint_open"},
    )
    terminate()


def addin_ready(service) -> bool:
    """Is the iSpring tab on the ribbon of the PowerPoint we just opened?

    Checked before every publish, because Office quietly disables the add-in
    after any hard close and the tab is the only proof it actually loaded.
    A missing tab is fixed in two steps: switch it on inside the running
    PowerPoint, and if that does not work, put the registry right and
    restart PowerPoint.
    """
    from powerpoint.addin import connect_addin
    from publisher.ispring_cloud import find_powerpoint_window, has_addin_tab

    try:
        window = find_powerpoint_window()
    except Exception as exc:  # noqa: BLE001
        logger.warning(
            "could not find the PowerPoint window to check the add-in",
            extra={"stage": "powerpoint_open", "detail": str(exc)},
        )
        return False

    from publisher.ispring_cloud import ribbon_tabs

    if has_addin_tab(window):
        return True

    logger.warning(
        "the iSpring Suite tab is not on the ribbon; switching the add-in on",
        extra={"stage": "powerpoint_open", "tabs": ribbon_tabs(window)},
    )
    if connect_addin(getattr(service, "app", None)) and has_addin_tab(window, timeout_s=30):
        logger.info(
            "iSpring tab came back without restarting PowerPoint",
            extra={"stage": "powerpoint_open"},
        )
        return True
    return False


def open_with_repair(pptx: Path, settings: Settings):
    """Open the deck through PowerPoint, repairing it if need be.

    Uses the service's OpenAndRepair path, with a watcher answering the repair
    prompt from another thread — PowerPoint's COM open does not return while
    that dialog is up.

    Returns ``(service, path_in_use, prompts_answered)``. A repaired deck is
    saved under a new name, so the path in use is not always the one passed in,
    and PowerPoint sometimes repairs silently — the prompt count is the other
    way of knowing it happened.
    """
    from powerpoint.addin import ensure_addin_enabled
    from powerpoint.service import PowerPointService, terminate_powerpoint_processes
    from publisher.ispring_cloud import ensure_com

    ensure_com()

    ensure_usable_powerpoint(terminate_powerpoint_processes)

    # With PowerPoint closed, put the add-in back if Office disabled it.
    # Doing this while PowerPoint runs achieves nothing: it rewrites these
    # keys when it exits.
    if not ensure_addin_enabled() and not powerpoint_is_running():
        logger.warning(
            "continuing without a confirmed iSpring add-in; the publish will "
            "fail if the ribbon tab does not appear",
            extra={"stage": "powerpoint_open"},
        )

    last_error: Exception | None = None
    for attempt in (1, 2, 3):
        service = PowerPointService(settings)
        try:
            service.start()
            with RepairPromptWatcher() as watcher:
                service.open(pptx, timeout_s=settings.powerpoint_open_timeout_seconds)
        except PowerPointAutomationError as exc:
            last_error = exc
            try:
                service.quit()
            except Exception:  # noqa: BLE001
                pass
            terminate_powerpoint_processes()
            if attempt < 3:
                logger.warning(
                    "PowerPoint would not open the deck; retrying with a fresh copy",
                    extra={"stage": "powerpoint_open", "detail": str(exc)},
                )
                time.sleep(3)
                continue
            raise
        if watcher.answered:
            logger.info(
                "PowerPoint repaired the deck",
                extra={"stage": "powerpoint_open", "prompts_answered": watcher.answered},
            )

        # The deck is open. Nothing can be published without the iSpring tab,
        # so check it now and put the add-in back before going further.
        if addin_ready(service):
            in_use = service.current_pptx or pptx
            return service, Path(in_use), watcher.answered
        if attempt == 3:
            logger.error(
                "still no iSpring Suite tab after two repairs; publishing anyway "
                "with whatever is on the ribbon. If this machine shows the "
                "iSpring Free ribbon, the Suite licence is not active on it",
                extra={"stage": "powerpoint_open"},
            )
            in_use = service.current_pptx or pptx
            return service, Path(in_use), watcher.answered

        logger.warning(
            "restarting PowerPoint to load the iSpring add-in",
            extra={"stage": "powerpoint_open", "attempt": attempt},
        )
        try:
            service.quit()
        except Exception:  # noqa: BLE001
            pass
        terminate_powerpoint_processes()
        time.sleep(3)
        # This is exactly what scripts/fix_ispring_addin.py --fix does, and it
        # only works with PowerPoint closed: Office writes these keys back as
        # it exits, so repairing them under a running PowerPoint does nothing.
        fixed = ensure_addin_enabled()
        logger.info(
            "ran the add-in repair",
            extra={"stage": "powerpoint_open", "attempt": attempt, "ok": fixed},
        )
        time.sleep(2)

    raise last_error  # unreachable; the loop either returns or raises


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
    job_id: int | str = "",
    content_name: str = "",
    keep_files: bool = False,
) -> CloudJobResult:
    """Download, repair if needed, publish, and return the iframe URL.

    The deck is published under the queue row's id, not its name. Two
    materials can share a name, and the browser half then has no way to tell
    which row is the one just published. job_id is unique per queue row. A
    later retry of the same row reuses that job_id and overwrites the existing
    presentation. The name a person should see is set on the cover instead.
    """
    publish_name = str(content_name or job_id or material_id).strip()
    if not url and not pptx:
        raise ValueError("pass either url or pptx")
    started = time.monotonic()
    workspace = Path(settings.temp_root) / f"cloud-{publish_name}"
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
        service, in_use, prompts = open_with_repair(deck, settings)
        repaired = damaged or in_use != deck or prompts > 0

        def publish_once() -> CloudPublishResult:
            return publish_to_cloud(
                in_use,
                institution=institution_name,
                content_name=publish_name,
                cover_title=material_name,
                parent_folders=settings.ispring_parent_folders,
                cdp_url=settings.ispring_chrome_cdp_url,
                publish_timeout_s=settings.ispring_publish_timeout_seconds,
                close_powerpoint_after=False,
                skip_open=True,
                browser_profile_dir=settings.ispring_chrome_profile_dir,
                browser_path=settings.ispring_chrome_path,
                cloud_url=settings.ispring_cloud_url,
            )

        try:
            result: CloudPublishResult = publish_once()
        except ProjectMissingError:
            # A new institution: nobody has made it a folder yet. Make one and
            # publish again. Once only — if the folder is still not in the
            # picker after that, something is wrong that another round of the
            # same will not fix.
            logger.info(
                "no folder for this institution yet; creating one",
                extra={
                    "stage": "ispring_publish",
                    "institution": institution_name,
                    "parent": settings.ispring_new_institution_parent,
                },
            )
            ensure_project_folder(
                institution_name,
                settings.ispring_new_institution_parent,
                settings.ispring_chrome_cdp_url,
                profile_dir=settings.ispring_chrome_profile_dir,
                chrome_path=settings.ispring_chrome_path,
                cloud_url=settings.ispring_cloud_url,
            )
            result = publish_once()
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
        content_name=publish_name,
        source_name=source_name,
        repaired=repaired,
        elapsed_s=time.monotonic() - started,
    )
