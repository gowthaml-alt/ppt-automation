"""One-shot local run: give a PPTX, upload HTML5, verify the page loads.

Does not poll Node.js. Used on a Windows laptop to watch PowerPoint and to
prove upload + page check. Real iSpring is still optional via --use-fake-ispring.
"""

from __future__ import annotations

import logging
import sys
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urlsplit

from api_client.jobs import Job
from browser_test.local_serve import LocalOutputChecker
from config.settings import Settings
from downloader.local import copy_local_pptx
from powerpoint.service import NoOpPowerPointService, PowerPointService
from publisher.ispring import FakePublisher, get_publisher
from storage.backend import get_storage_backend
from utils.cleanup import CleanupService
from utils.exceptions import PptAutomationError
from utils.paths import JobPaths
from worker.pipeline import Pipeline

logger = logging.getLogger(__name__)


@dataclass
class LocalRunResult:
    ok: bool
    error: str | None = None
    iframe_url: str | None = None
    local_index: Path | None = None
    checked_url: str | None = None
    browser_ok: bool = False


class SilentBackend:
    """Records callback payloads but does not call Node."""

    def __init__(self) -> None:
        self.results: list[dict] = []

    def send_job_result(
        self,
        job: Job,
        *,
        status: str,
        iframe_url: str | None = None,
        error_message: str | None = None,
    ) -> None:
        self.results.append(
            {
                "queue_id": job.queue_id,
                "material_id": job.material_id,
                "status": status,
                "iframe_url": iframe_url,
                "error_message": error_message,
            }
        )
        logger.info(
            "local run skipped Node callback",
            extra={"status": status, "stage": "callback"},
        )


class LocalJobBrowser:
    def __init__(
        self, settings: Settings, *, use_playwright: bool = False, timeout_s: float = 15
    ) -> None:
        self._settings = settings
        self._use_playwright = use_playwright
        self._timeout_s = timeout_s
        self.checked_url = ""
        self.ok = False

    def check(self, iframe_url: str) -> None:
        if self._settings.storage_backend == "local_fs":
            folder = local_output_dir(iframe_url, self._settings)
            self.checked_url = LocalOutputChecker(
                timeout_s=self._timeout_s, use_playwright=self._use_playwright
            ).serve_and_check(folder)
        else:
            from browser_test.playwright_check import PlaywrightChecker

            PlaywrightChecker(self._settings).check(iframe_url)
            self.checked_url = iframe_url
        self.ok = True


def local_output_dir(iframe_url: str, settings: Settings) -> Path:
    path = urlsplit(iframe_url).path
    public = urlsplit(settings.local_storage_public_base_url or "").path.rstrip("/")
    relative = path
    if public and path.startswith(public):
        relative = path[len(public) :]
    relative = relative.lstrip("/")
    parent = Path(relative).parent
    return Path(settings.local_storage_root) / parent


def run_local_job(
    pptx: Path,
    *,
    settings: Settings,
    queue_id: int = 1,
    material_id: int = 1,
    material_name: str = "local-test",
    institution_name: str | None = "local",
    skip_powerpoint: bool = False,
    use_fake_ispring: bool = True,
    skip_callback: bool = True,
    use_playwright: bool = False,
) -> LocalRunResult:
    source = Path(pptx).expanduser().resolve()
    job = Job(
        queue_id=queue_id,
        material_id=material_id,
        ppt_file_url=str(source),
        material_name=material_name,
        institution_name=institution_name,
    )

    if skip_powerpoint or sys.platform != "win32":
        powerpoint = NoOpPowerPointService()
    else:
        powerpoint = PowerPointService(settings)

    publisher = FakePublisher() if use_fake_ispring else get_publisher(settings)
    storage = get_storage_backend(settings)
    if skip_callback:
        backend = SilentBackend()
    else:
        from api_client.jobs import BackendClient

        backend = BackendClient(settings)
    browser = LocalJobBrowser(
        settings,
        use_playwright=use_playwright,
        timeout_s=float(settings.browser_test_timeout_seconds),
    )

    def _download(url: str, paths: JobPaths, job_settings: Settings, **kwargs):
        return copy_local_pptx(source, paths)

    pipeline = Pipeline(
        settings,
        backend=backend,
        powerpoint=powerpoint,
        publisher=publisher,
        storage=storage,
        browser=browser,
        cleanup=CleanupService(settings),
        downloader=_download,
    )

    if use_fake_ispring:
        logger.warning(
            "local run uses fake iSpring output (fixture HTML). "
            "This tests upload + page load, not a real iSpring conversion.",
            extra={"stage": "ispring_publish"},
        )

    try:
        pipeline.process(job)
    except PptAutomationError as exc:
        return LocalRunResult(ok=False, error=exc.user_message or str(exc))

    recorded = getattr(backend, "results", None) or []
    iframe_url = recorded[-1].get("iframe_url") if recorded else None
    local_index = None
    if settings.storage_backend == "local_fs":
        from storage.backend import output_prefix

        prefix = output_prefix(material_id, settings)
        candidate = Path(settings.local_storage_root) / prefix / "index.html"
        if candidate.is_file():
            local_index = candidate
        elif iframe_url:
            local_index = local_output_dir(iframe_url, settings) / "index.html"

    return LocalRunResult(
        ok=True,
        iframe_url=iframe_url,
        local_index=local_index,
        checked_url=browser.checked_url or None,
        browser_ok=browser.ok,
    )
