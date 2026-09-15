# Processing Pipeline Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Install the Windows processing pipeline behind the existing `JobPipeline` protocol so a claimed job is downloaded, published through PowerPoint/iSpring adapters, uploaded, reported to the existing backend, and cleaned up — without claiming the end-to-end automation works until a real Windows publish is tested.

**Architecture:** `JobRunner` orchestrates named stages on a single `ThreadPoolExecutor(max_workers=1)` so COM work never blocks the asyncio loop. Every environment-specific boundary is a constructor-injected protocol (`ISpringPublisher`, `StorageBackend`, `BackendClient`, PowerPoint service, Slack, cleanup). The default iSpring adapter is `not_configured` and fails loudly; the real publishing mechanism is completed only after `scripts/probe_ispring.py` is run on the Cloud PC.

**Tech Stack:** Python 3.11+, FastAPI (already present), httpx, boto3, psutil, SQLAlchemy Core repository (already present), pywin32/pywinauto on Windows only.

## Global Constraints

- Design spec: `docs/superpowers/specs/2026-09-14-ppt-automation-worker-design.md`. Read it before starting.
- Intake plan already implemented: `docs/superpowers/plans/2026-09-14-job-intake-service.md`. Do not re-implement claim, auth, schemas, or the job table access layer.
- The entire test suite MUST run and pass on macOS and Linux with no PowerPoint, no iSpring, and no MySQL installed.
- No module in `app/api/` or `app/services/` may import a Windows-only module (`win32com`, `pythoncom`, `pywintypes`, `pywinauto`) at module scope. COM goes through `app/integrations/windows_com.py`.
- Write NO DDL. Never reference `current_step`, `worker_type`, `retry_count`, `retry_at`, or `iframe_url` as columns on `ppt_automation_jobs`.
- Status values: `0` PENDING, `1` PROCESSING, `2` COMPLETED, `3` FAILED.
- `COMPLETED` is written only after the backend update returns success.
- Do not fabricate an iSpring CLI, COM API, ProgID, or control identifier. Stubs raise with instructions pointing at the probe script.
- No bare `except:`. No `except Exception` without `exc_info=True` on the log call. The only designed swallow is Slack.
- Never log an API key, password, token, webhook URL, or a URL's query string. Use `app.core.logging_config.scrub_url`.
- Every filesystem path is derived from the integer `job_id`. The downloaded file is always `input/source.pptx`.
- Reuse existing helpers: `build_job_paths`, `create_job_dirs`, `assert_within`, `free_disk_gb`, `iter_files`, `validate_download_url`, `assert_looks_like_pptx`, `JobRow`, `JobStatus`, `ERROR_MESSAGE_MAX_CHARS`.
- This workspace may not have a git repository. If `git status` fails, skip commit steps rather than running `git init` unless the operator asked for a repo.
- Do not claim the complete automation works. The exit criterion for that claim is `docs/MANUAL_WINDOWS_TESTS.md`.

## File structure

| File | Responsibility |
| --- | --- |
| `app/integrations/http_client.py` | Shared sync httpx client with connect/read timeouts |
| `app/services/download_service.py` | Stream PPT to `.part`, rename, validate |
| `app/utils/output_validation.py` | Local iSpring folder checks + manifest |
| `app/services/cleanup_service.py` | `finally`-safe job directory deletion |
| `app/utils/process_utils.py` | Process listing / optional orphan kill via psutil |
| `app/integrations/windows_com.py` | Platform-guarded COM apartment + Dispatch + error translation |
| `app/services/powerpoint_service.py` | Start / open / close / quit / availability |
| `app/services/ispring_service.py` | Publisher protocol, factory, `not_configured` / `fake` / stubs |
| `app/services/storage_service.py` | `UploadResult`, MIME map, storage factory |
| `app/integrations/storage_backends/local_fs.py` | Dev/test copy backend |
| `app/integrations/storage_backends/s3.py` | S3-compatible upload + verify |
| `app/integrations/storage_backends/http_api.py` | Documented stub |
| `app/services/backend_service.py` | Notify existing backend; record orphaned outputs |
| `app/services/slack_service.py` | Failure webhook; never raises |
| `app/services/job_runner.py` | Stage orchestration, status writes, COM thread |
| `scripts/probe_ispring.py` | Windows fact-gatherer; no invented APIs |
| `tests/fixtures/ispring_output/` | Minimal fake HTML5 package |
| `README.md` | Operator documentation |
| `docs/MANUAL_WINDOWS_TESTS.md` | Windows-only integration steps |

Existing files modified only in the wiring tasks: `app/main.py`, `app/services/health_service.py` (remove the "not wired up yet" probes once real ones exist), `tests/fakes.py`.

---

### Task 1: HTTP client and streaming download

**Files:**
- Create: `app/integrations/http_client.py`
- Create: `app/services/download_service.py`
- Test: `tests/test_download_service.py`

**Interfaces:**
- Consumes: `Settings.download_timeout_seconds`, `Settings.max_ppt_size_bytes`, `Settings.download_allowed_hosts`, `JobPaths.source_pptx`, `JobPaths.partial_pptx`, `validate_download_url`, `assert_looks_like_pptx`, `scrub_url`, `DownloadError`
- Produces: `create_sync_client(*, timeout_s: float, connect_s: float = 10.0) -> httpx.Client`, `download_pptx(url: str, paths: JobPaths, settings: Settings, *, client: httpx.Client | None = None) -> Path`

- [ ] **Step 1: Write the failing test `tests/test_download_service.py`**

```python
from pathlib import Path

import httpx
import pytest

from app.core.config import Settings
from app.exceptions import DownloadError
from app.services.download_service import download_pptx
from app.utils.file_utils import build_job_paths, create_job_dirs

PPTX = b"PK\x03\x04" + b"payload-bytes"


def _settings(tmp_path: Path, **overrides) -> Settings:
    values = {
        "app_env": "development",
        "worker_api_key": "test-key",
        "temp_root": str(tmp_path / "jobs"),
        "log_root": str(tmp_path / "logs"),
        "max_ppt_size_mb": 1,
        "download_timeout_seconds": 5,
        "download_allowed_hosts": [],
    }
    values.update(overrides)
    return Settings(**values)


def _paths(tmp_path: Path):
    paths = build_job_paths(tmp_path / "jobs", 101)
    create_job_dirs(paths)
    return paths


def _client(handler) -> httpx.Client:
    return httpx.Client(transport=httpx.MockTransport(handler))


def test_downloads_to_source_pptx_and_removes_part(tmp_path):
    def handler(request: httpx.Request) -> httpx.Response:
        assert request.method == "GET"
        return httpx.Response(200, content=PPTX)

    paths = _paths(tmp_path)
    result = download_pptx(
        "https://storage.example.com/input/biology.pptx",
        paths,
        _settings(tmp_path),
        client=_client(handler),
    )

    assert result == paths.source_pptx
    assert result.read_bytes() == PPTX
    assert not paths.partial_pptx.exists()


def test_rejects_non_http_scheme(tmp_path):
    paths = _paths(tmp_path)
    with pytest.raises(DownloadError) as exc:
        download_pptx("file:///tmp/x.pptx", paths, _settings(tmp_path))
    assert exc.value.stage == "download"


def test_http_error_is_download_error(tmp_path):
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(404, content=b"missing")

    paths = _paths(tmp_path)
    with pytest.raises(DownloadError):
        download_pptx(
            "https://storage.example.com/missing.pptx",
            paths,
            _settings(tmp_path),
            client=_client(handler),
        )
    assert not paths.source_pptx.exists()
    assert not paths.partial_pptx.exists()


def test_timeout_is_download_error(tmp_path):
    def handler(request: httpx.Request) -> httpx.Response:
        raise httpx.ReadTimeout("slow")

    paths = _paths(tmp_path)
    with pytest.raises(DownloadError) as exc:
        download_pptx(
            "https://storage.example.com/slow.pptx",
            paths,
            _settings(tmp_path),
            client=_client(handler),
        )
    assert "timeout" in str(exc.value).lower()


def test_oversize_aborts_before_rename(tmp_path):
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(200, content=b"PK\x03\x04" + b"x" * (2 * 1024 * 1024))

    paths = _paths(tmp_path)
    with pytest.raises(DownloadError) as exc:
        download_pptx(
            "https://storage.example.com/huge.pptx",
            paths,
            _settings(tmp_path, max_ppt_size_mb=1),
            client=_client(handler),
        )
    assert "size" in str(exc.value).lower()
    assert not paths.source_pptx.exists()
    assert not paths.partial_pptx.exists()


def test_html_error_page_fails_magic_bytes(tmp_path):
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(200, content=b"<html>not a pptx</html>")

    paths = _paths(tmp_path)
    with pytest.raises(DownloadError) as exc:
        download_pptx(
            "https://storage.example.com/fake.pptx",
            paths,
            _settings(tmp_path),
            client=_client(handler),
        )
    assert "valid PowerPoint" in exc.value.user_message


def test_host_allowlist_is_enforced(tmp_path):
    paths = _paths(tmp_path)
    with pytest.raises(DownloadError):
        download_pptx(
            "https://evil.example.com/x.pptx",
            paths,
            _settings(tmp_path, download_allowed_hosts=["storage.example.com"]),
        )
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pytest tests/test_download_service.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'app.services.download_service'`

- [ ] **Step 3: Write `app/integrations/http_client.py`**

```python
"""Shared synchronous HTTP client. Timeouts are always explicit."""

from __future__ import annotations

import httpx


def create_sync_client(*, timeout_s: float, connect_s: float = 10.0) -> httpx.Client:
    """Build a client with a total budget and a shorter connect budget.

    ``timeout_s`` is the read/overall budget from configuration.
    ``connect_s`` stays small so a black-holed host fails fast.
    """
    return httpx.Client(
        timeout=httpx.Timeout(timeout_s, connect=min(connect_s, timeout_s)),
        follow_redirects=True,
        headers={"User-Agent": "ppt-automation-worker/1.0"},
    )
```

- [ ] **Step 4: Write `app/services/download_service.py`**

```python
"""Stream an input PPT to the job directory. Never trusts the remote filename."""

from __future__ import annotations

import logging
from pathlib import Path

import httpx

from app.core.config import Settings
from app.core.logging_config import scrub_url
from app.exceptions import DownloadError
from app.integrations.http_client import create_sync_client
from app.utils.file_utils import JobPaths
from app.utils.validators import assert_looks_like_pptx, validate_download_url

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
                    if written > max_bytes:
                        raise DownloadError(
                            f"download exceeded max size {max_bytes} bytes "
                            f"({scrub_url(url)})",
                            user_message="The PowerPoint file is larger than the configured limit.",
                        )
                    handle.write(chunk)
    except DownloadError:
        raise
    except httpx.HTTPError:
        raise


def _remove_if_exists(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except OSError:
        logger.warning("could not remove %s", path.name, exc_info=True)
```

- [ ] **Step 5: Run the test to verify it passes**

Run: `pytest tests/test_download_service.py -v`
Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add app/integrations/http_client.py app/services/download_service.py tests/test_download_service.py
git commit -m "feat: stream PPT downloads to a job-scoped source.pptx"
```

---

### Task 2: Output validation and cleanup

**Files:**
- Create: `app/utils/output_validation.py`
- Create: `app/services/cleanup_service.py`
- Test: `tests/test_output_validation.py`
- Test: `tests/test_cleanup_service.py`

**Interfaces:**
- Consumes: `JobPaths`, `iter_files`, `CleanupError`, `OutputValidationError`, `Settings.keep_failed_job_files`, `Settings.log_root`
- Produces: `validate_ispring_output(output_dir: Path) -> list[dict]`, `CleanupService.run(paths: JobPaths, *, success: bool, original_error: BaseException | None = None) -> None`

- [ ] **Step 1: Write the failing tests**

`tests/test_output_validation.py`:

```python
from pathlib import Path

import pytest

from app.exceptions import OutputValidationError
from app.utils.output_validation import validate_ispring_output


def _package(root: Path) -> Path:
    (root / "data").mkdir()
    (root / "index.html").write_text("<html>ok</html>", encoding="utf-8")
    (root / "data" / "app.js").write_text("console.log(1)", encoding="utf-8")
    return root


def test_valid_package_returns_manifest(tmp_path):
    output = _package(tmp_path / "output")
    manifest = validate_ispring_output(output)
    names = {row["path"] for row in manifest}
    assert "index.html" in names
    assert "data/app.js" in names
    assert all(row["size"] > 0 for row in manifest)


def test_missing_index_is_error(tmp_path):
    output = tmp_path / "output"
    output.mkdir()
    (output / "data").mkdir()
    (output / "data" / "app.js").write_text("x", encoding="utf-8")
    with pytest.raises(OutputValidationError):
        validate_ispring_output(output)


def test_empty_index_is_error(tmp_path):
    output = tmp_path / "output"
    (output / "data").mkdir(parents=True)
    (output / "index.html").write_text("", encoding="utf-8")
    (output / "data" / "app.js").write_text("x", encoding="utf-8")
    with pytest.raises(OutputValidationError):
        validate_ispring_output(output)


def test_no_asset_subdirectory_is_error(tmp_path):
    output = tmp_path / "output"
    output.mkdir()
    (output / "index.html").write_text("<html>stub</html>", encoding="utf-8")
    with pytest.raises(OutputValidationError):
        validate_ispring_output(output)
```

`tests/test_cleanup_service.py`:

```python
from pathlib import Path

from app.core.config import Settings
from app.services.cleanup_service import CleanupService
from app.utils.file_utils import build_job_paths, create_job_dirs


def _settings(tmp_path: Path, **overrides) -> Settings:
    values = {
        "app_env": "development",
        "worker_api_key": "test-key",
        "temp_root": str(tmp_path / "jobs"),
        "log_root": str(tmp_path / "logs"),
        "keep_failed_job_files": False,
    }
    values.update(overrides)
    return Settings(**values)


def _job_with_files(tmp_path: Path):
    paths = build_job_paths(tmp_path / "jobs", 101)
    create_job_dirs(paths)
    paths.source_pptx.write_bytes(b"PK\x03\x04xx")
    (paths.log_dir / "job.log").write_text("ran", encoding="utf-8")
    (paths.output_dir / "index.html").write_text("x", encoding="utf-8")
    return paths


def test_success_deletes_the_job_directory(tmp_path):
    paths = _job_with_files(tmp_path)
    CleanupService(_settings(tmp_path)).run(paths, success=True)
    assert not paths.root.exists()


def test_failure_copies_logs_then_deletes(tmp_path):
    paths = _job_with_files(tmp_path)
    settings = _settings(tmp_path)
    CleanupService(settings).run(
        paths, success=False, original_error=RuntimeError("publish failed")
    )
    preserved = Path(settings.log_root) / "jobs" / "101" / "job.log"
    assert preserved.read_text(encoding="utf-8") == "ran"
    assert not paths.root.exists()


def test_keep_failed_job_files_preserves_directory(tmp_path):
    paths = _job_with_files(tmp_path)
    settings = _settings(tmp_path, keep_failed_job_files=True)
    CleanupService(settings).run(paths, success=False)
    assert paths.source_pptx.exists()
    assert paths.root.exists()


def test_cleanup_error_does_not_replace_original(tmp_path, monkeypatch):
    import shutil

    paths = _job_with_files(tmp_path)
    monkeypatch.setattr(
        shutil, "rmtree", lambda *args, **kwargs: (_ for _ in ()).throw(OSError("locked"))
    )
    # Must not raise: cleanup failures are logged, not propagated, when an
    # original_error is supplied. When called without one, they still must
    # not crash the pipeline — run() never raises.
    CleanupService(_settings(tmp_path)).run(
        paths, success=True, original_error=RuntimeError("upload failed")
    )
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `pytest tests/test_output_validation.py tests/test_cleanup_service.py -v`
Expected: FAIL with `ModuleNotFoundError`

- [ ] **Step 3: Write `app/utils/output_validation.py`**

```python
"""Local checks on a published iSpring HTML5 folder."""

from __future__ import annotations

import logging
from pathlib import Path

from app.exceptions import OutputValidationError
from app.utils.file_utils import iter_files

logger = logging.getLogger(__name__)


def validate_ispring_output(output_dir: Path) -> list[dict]:
    if not output_dir.exists() or not output_dir.is_dir():
        raise OutputValidationError(
            f"publish output directory is missing: {output_dir}",
            user_message="iSpring did not produce an output folder.",
        )

    manifest = [
        {"path": relative, "size": absolute.stat().st_size}
        for absolute, relative in iter_files(output_dir)
    ]
    if not manifest:
        raise OutputValidationError(
            f"publish output directory is empty: {output_dir}",
            user_message="iSpring produced an empty output folder.",
        )

    index = output_dir / "index.html"
    if not index.exists() or index.stat().st_size == 0:
        raise OutputValidationError(
            "publish output is missing a non-empty index.html",
            user_message="iSpring output is missing index.html.",
        )

    has_asset_dir = any(
        (output_dir / relative.split("/")[0]).is_dir()
        for relative in (row["path"] for row in manifest)
        if "/" in row["path"]
    )
    if not has_asset_dir:
        raise OutputValidationError(
            "publish output has no asset subdirectory; refusing a stub package",
            user_message="iSpring output looks incomplete.",
        )

    logger.info(
        "output validated",
        extra={"file_count": len(manifest), "stage": "output_validate"},
    )
    return manifest
```

- [ ] **Step 4: Write `app/services/cleanup_service.py`**

```python
"""Delete job-scoped temporary files. Always invoked from a finally block."""

from __future__ import annotations

import logging
import shutil
from pathlib import Path

from app.core.config import Settings
from app.exceptions import CleanupError
from app.utils.file_utils import JobPaths

logger = logging.getLogger(__name__)


class CleanupService:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings

    def run(
        self,
        paths: JobPaths,
        *,
        success: bool,
        original_error: BaseException | None = None,
    ) -> None:
        """Best-effort cleanup. Never raises.

        A cleanup failure is logged as its own event and never replaces
        ``original_error``. The caller records the original error first.
        """
        try:
            self._run(paths, success=success)
        except Exception as exc:
            logger.error(
                "cleanup failed",
                extra={"stage": "cleanup", "success": success},
                exc_info=True,
            )
            if original_error is None:
                # Surface only when cleanup itself is the only failure.
                logger.error(
                    "cleanup failed with no prior processing error: %s",
                    CleanupError(str(exc), user_message="Temporary files could not be deleted."),
                    exc_info=True,
                )

    def _run(self, paths: JobPaths, *, success: bool) -> None:
        if not paths.root.exists():
            return
        if success:
            shutil.rmtree(paths.root)
            logger.info("job directory deleted", extra={"stage": "cleanup"})
            return
        self._preserve_logs(paths)
        if self._settings.keep_failed_job_files:
            logger.info(
                "KEEP_FAILED_JOB_FILES is set; leaving the job directory",
                extra={"stage": "cleanup"},
            )
            return
        shutil.rmtree(paths.root)
        logger.info("failed-job directory deleted after log copy", extra={"stage": "cleanup"})

    def _preserve_logs(self, paths: JobPaths) -> None:
        if not paths.log_dir.exists():
            return
        destination = Path(self._settings.log_root) / "jobs" / paths.root.name
        destination.mkdir(parents=True, exist_ok=True)
        for item in paths.log_dir.iterdir():
            if item.is_file():
                shutil.copy2(item, destination / item.name)
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `pytest tests/test_output_validation.py tests/test_cleanup_service.py -v`
Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add app/utils/output_validation.py app/services/cleanup_service.py tests/test_output_validation.py tests/test_cleanup_service.py
git commit -m "feat: validate iSpring output and clean up job directories"
```

---

### Task 3: Process helpers, COM boundary, and PowerPoint service

**Files:**
- Create: `app/utils/process_utils.py`
- Create: `app/integrations/windows_com.py`
- Create: `app/services/powerpoint_service.py`
- Test: `tests/test_process_utils.py`
- Test: `tests/test_windows_com.py`
- Test: `tests/test_powerpoint_service.py`
- Modify: `tests/fakes.py` (add `FakePowerPointService`)

**Interfaces:**
- Consumes: `PowerPointAutomationError`, `Settings.powerpoint_exe_path`, `Settings.powerpoint_kill_orphans`, `Settings.powerpoint_orphan_max_age_seconds`, `Settings.powerpoint_start_timeout_seconds`, `Settings.powerpoint_open_timeout_seconds`, `CheckResult`
- Produces: `list_processes(name: str) -> list[ProcessInfo]`, `terminate_orphans(name: str, *, max_age_s: int) -> int`, `com_apartment()`, `dispatch(prog_id: str)`, `translate_com_error(exc: BaseException) -> PowerPointAutomationError`, `PowerPointService` with `is_available() -> CheckResult`, `start() -> None`, `open(path: Path) -> None`, `close() -> None`, `quit() -> None`, `cleanup_orphans() -> int`

- [ ] **Step 1: Write the failing tests**

`tests/test_process_utils.py`:

```python
import types

from app.utils import process_utils


class _Proc:
    def __init__(self, pid, name, age, cmdline=None):
        self.info = {"pid": pid, "name": name, "create_time": 0}
        self._name = name
        self._age = age
        self._cmdline = cmdline or [name]
        self.killed = False

    def name(self):
        return self._name

    def create_time(self):
        return 1_000_000

    def cmdline(self):
        return self._cmdline

    def kill(self):
        self.killed = True


def test_list_processes_filters_by_name(monkeypatch):
    procs = [_Proc(1, "POWERPNT.EXE", 10), _Proc(2, "excel.exe", 10)]

    def fake_iter(attrs):
        return procs

    monkeypatch.setattr(process_utils.psutil, "process_iter", fake_iter)
    monkeypatch.setattr(process_utils.time, "time", lambda: 1_000_010)
    found = process_utils.list_processes("powerpnt.exe")
    assert [p.pid for p in found] == [1]


def test_terminate_orphans_respects_age(monkeypatch):
    young = _Proc(1, "POWERPNT.EXE", 10)
    old = _Proc(2, "POWERPNT.EXE", 10)

    def fake_iter(attrs):
        return [young, old]

    monkeypatch.setattr(process_utils.psutil, "process_iter", fake_iter)
    # young.create_time is 1_000_000; "now" is 1_000_010 => age 10s
    # old we will patch after listing by making create_time older via side effect
    monkeypatch.setattr(process_utils.time, "time", lambda: 1_000_010)

    # Both are 10 seconds old; threshold 60 => none killed
    killed = process_utils.terminate_orphans("powerpnt.exe", max_age_s=60)
    assert killed == 0
    assert not young.killed

    monkeypatch.setattr(process_utils.time, "time", lambda: 1_003_700)
    killed = process_utils.terminate_orphans("powerpnt.exe", max_age_s=60)
    assert killed == 2
```

`tests/test_windows_com.py`:

```python
import sys
from types import SimpleNamespace

from app.exceptions import PowerPointAutomationError
from app.integrations import windows_com


def test_translate_com_error_extracts_hresult():
    exc = SimpleNamespace(args=((0x80040154, "Class not registered"),))
    wrapped = windows_com.translate_com_error(exc)
    assert isinstance(wrapped, PowerPointAutomationError)
    assert "0x80040154" in wrapped.message or "Class not registered" in wrapped.message


def test_com_apartment_is_a_noop_off_windows(monkeypatch):
    monkeypatch.setattr(windows_com, "sys", SimpleNamespace(platform="darwin"))
    with windows_com.com_apartment():
        pass


def test_dispatch_raises_off_windows(monkeypatch):
    monkeypatch.setattr(sys, "platform", "darwin")
    try:
        windows_com.dispatch("PowerPoint.Application")
        raised = False
    except PowerPointAutomationError:
        raised = True
    assert raised
```

`tests/test_powerpoint_service.py`:

```python
from pathlib import Path

import pytest

from app.core.config import Settings
from app.exceptions import PowerPointAutomationError
from app.services.powerpoint_service import PowerPointService


def _settings(tmp_path: Path, **overrides) -> Settings:
    values = {
        "app_env": "development",
        "worker_api_key": "test-key",
        "temp_root": str(tmp_path / "jobs"),
        "log_root": str(tmp_path / "logs"),
        "powerpoint_exe_path": "",
        "powerpoint_kill_orphans": False,
    }
    values.update(overrides)
    return Settings(**values)


def test_is_available_is_false_off_windows(tmp_path, monkeypatch):
    monkeypatch.setattr("app.services.powerpoint_service.sys.platform", "darwin")
    result = PowerPointService(_settings(tmp_path)).is_available()
    assert result.name == "powerpoint"
    assert result.ok is False
    assert "Windows" in result.detail


def test_start_raises_off_windows(tmp_path, monkeypatch):
    monkeypatch.setattr("app.services.powerpoint_service.sys.platform", "darwin")
    with pytest.raises(PowerPointAutomationError):
        PowerPointService(_settings(tmp_path)).start()


def test_close_and_quit_are_idempotent_when_never_started(tmp_path):
    service = PowerPointService(_settings(tmp_path))
    service.close()
    service.quit()
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `pytest tests/test_process_utils.py tests/test_windows_com.py tests/test_powerpoint_service.py -v`
Expected: FAIL with import errors

- [ ] **Step 3: Write `app/utils/process_utils.py`**

```python
"""Process inspection. Used to find orphaned POWERPNT.exe processes."""

from __future__ import annotations

import logging
import time
from dataclasses import dataclass

import psutil

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class ProcessInfo:
    pid: int
    name: str
    age_seconds: float


def list_processes(name: str) -> list[ProcessInfo]:
    wanted = name.lower()
    now = time.time()
    found: list[ProcessInfo] = []
    for proc in psutil.process_iter(["pid", "name", "create_time"]):
        try:
            proc_name = proc.name()
        except (psutil.Error, OSError):
            continue
        if proc_name.lower() != wanted:
            continue
        try:
            age = now - proc.create_time()
        except (psutil.Error, OSError):
            continue
        found.append(ProcessInfo(pid=proc.pid, name=proc_name, age_seconds=age))
    return found


def terminate_orphans(name: str, *, max_age_s: int) -> int:
    """Kill processes older than ``max_age_s``. Returns the number killed."""
    wanted = name.lower()
    now = time.time()
    killed = 0
    for proc in psutil.process_iter(["pid", "name", "create_time"]):
        try:
            if proc.name().lower() != wanted:
                continue
            if now - proc.create_time() < max_age_s:
                continue
            proc.kill()
            killed += 1
            logger.warning(
                "terminated orphaned process",
                extra={"pid": proc.pid, "name": name, "stage": "powerpoint"},
            )
        except (psutil.Error, OSError):
            logger.warning("could not inspect or kill process", exc_info=True)
    return killed
```

- [ ] **Step 4: Write `app/integrations/windows_com.py`**

```python
"""The only module allowed to import pywin32.

Imports are guarded so this file is importable on macOS and Linux. COM
objects never escape this module: callers receive plain Python objects
or ``PowerPointAutomationError``.
"""

from __future__ import annotations

import logging
import sys
from collections.abc import Iterator
from contextlib import contextmanager

from app.exceptions import PowerPointAutomationError

logger = logging.getLogger(__name__)


def _is_windows() -> bool:
    return sys.platform == "win32"


def translate_com_error(exc: BaseException) -> PowerPointAutomationError:
    hresult = None
    description = str(exc)
    args = getattr(exc, "args", ())
    if args:
        first = args[0]
        if isinstance(first, tuple) and first:
            hresult = first[0]
            if len(first) > 1:
                description = str(first[1])
        else:
            description = str(first)
    hex_code = f"0x{int(hresult) & 0xFFFFFFFF:08X}" if isinstance(hresult, int) else "unknown"
    return PowerPointAutomationError(
        f"COM error {hex_code}: {description}",
        user_message="PowerPoint automation failed.",
    )


@contextmanager
def com_apartment() -> Iterator[None]:
    """Pair CoInitialize / CoUninitialize on Windows. No-op elsewhere."""
    if not _is_windows():
        yield
        return
    try:
        import pythoncom  # type: ignore
    except ImportError as exc:
        raise PowerPointAutomationError(
            f"pywin32 is not installed: {exc}",
            user_message="PowerPoint automation libraries are not installed.",
        ) from exc
    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()


def dispatch(prog_id: str):
    if not _is_windows():
        raise PowerPointAutomationError(
            f"COM dispatch of {prog_id!r} is only available on Windows",
            user_message="PowerPoint is only available on the Windows worker.",
        )
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise PowerPointAutomationError(
            f"pywin32 is not installed: {exc}",
            user_message="PowerPoint automation libraries are not installed.",
        ) from exc
    try:
        return win32com.client.Dispatch(prog_id)
    except Exception as exc:  # noqa: BLE001 — translated immediately
        logger.error("COM dispatch failed", extra={"prog_id": prog_id}, exc_info=True)
        raise translate_com_error(exc) from exc
```

The `except Exception` above is the COM boundary. It MUST log with `exc_info=True` and MUST re-raise `PowerPointAutomationError`. Do not let `pywintypes` types escape.

- [ ] **Step 5: Write `app/services/powerpoint_service.py`**

```python
"""Desktop PowerPoint automation. No iSpring logic lives here."""

from __future__ import annotations

import logging
import sys
from pathlib import Path

from app.core.config import Settings
from app.exceptions import PowerPointAutomationError
from app.integrations import windows_com
from app.schemas.health import CheckResult
from app.utils.process_utils import terminate_orphans

logger = logging.getLogger(__name__)

POWERPOINT_PROCESS = "POWERPNT.EXE"
POWERPOINT_PROGID = "PowerPoint.Application"


class PowerPointService:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        self._app = None
        self._presentation = None

    def is_available(self) -> CheckResult:
        if sys.platform != "win32":
            return CheckResult(
                name="powerpoint",
                ok=False,
                detail="PowerPoint COM is only available on Windows",
            )
        configured = self._settings.powerpoint_exe_path
        if configured and not Path(configured).exists():
            return CheckResult(
                name="powerpoint",
                ok=False,
                detail="POWERPOINT_EXE_PATH does not exist",
            )
        # Registration check without launching: Dispatch of the ProgID is the
        # verified signal. A dedicated health probe may still be slow; keep it
        # to a path/platform check here and let a real start() fail loudly.
        return CheckResult(
            name="powerpoint",
            ok=True,
            detail="Windows host; COM will be opened at job time",
        )

    def start(self) -> None:
        if sys.platform != "win32":
            raise PowerPointAutomationError(
                "PowerPoint cannot be started off Windows",
                user_message="PowerPoint is only available on the Windows worker.",
            )
        logger.info("starting PowerPoint", extra={"stage": "powerpoint_open"})
        self._app = windows_com.dispatch(POWERPOINT_PROGID)
        try:
            self._app.DisplayAlerts = 1  # ppAlertsNone where the OM allows it
            self._app.Visible = True
        except Exception as exc:  # translated at the boundary if COM
            logger.warning("could not set PowerPoint automation flags", exc_info=True)
            if self._is_com(exc):
                raise windows_com.translate_com_error(exc) from exc

    def open(self, path: Path) -> None:
        if self._app is None:
            raise PowerPointAutomationError(
                "open() called before start()",
                user_message="PowerPoint was not started.",
            )
        resolved = Path(path).resolve()
        if not resolved.exists():
            raise PowerPointAutomationError(
                f"presentation does not exist: {resolved.name}",
                user_message="The PowerPoint file could not be opened.",
            )
        logger.info("opening presentation", extra={"stage": "powerpoint_open"})
        try:
            # WithWindow=True keeps iSpring's UI add-in in a real session.
            self._presentation = self._app.Presentations.Open(
                str(resolved), False, False, True
            )
        except Exception as exc:
            logger.error("Presentations.Open failed", exc_info=True)
            raise windows_com.translate_com_error(exc) from exc

    def close(self) -> None:
        presentation = self._presentation
        self._presentation = None
        if presentation is None:
            return
        try:
            presentation.Close()
        except Exception:
            logger.warning("presentation.Close failed", exc_info=True)

    def quit(self) -> None:
        self.close()
        app = self._app
        self._app = None
        if app is None:
            return
        try:
            app.Quit()
        except Exception:
            logger.warning("PowerPoint Quit failed", exc_info=True)

    def cleanup_orphans(self) -> int:
        if not self._settings.powerpoint_kill_orphans:
            return 0
        return terminate_orphans(
            POWERPOINT_PROCESS,
            max_age_s=self._settings.powerpoint_orphan_max_age_seconds,
        )

    @staticmethod
    def _is_com(exc: BaseException) -> bool:
        return exc.__class__.__module__.startswith("pywintypes") or exc.__class__.__name__ == "com_error"
```

- [ ] **Step 6: Append `FakePowerPointService` to `tests/fakes.py`**

Add imports only if needed. Append:

```python
@dataclass
class FakePowerPointService:
    available: bool = True
    started: bool = False
    opened: list[str] = field(default_factory=list)
    closed: int = 0
    quit_calls: int = 0
    fail_on: str | None = None

    def is_available(self):
        from app.schemas.health import CheckResult

        return CheckResult(
            name="powerpoint",
            ok=self.available,
            detail="fake",
        )

    def start(self) -> None:
        if self.fail_on == "start":
            from app.exceptions import PowerPointAutomationError

            raise PowerPointAutomationError("fake start failure")
        self.started = True

    def open(self, path) -> None:
        if self.fail_on == "open":
            from app.exceptions import PowerPointAutomationError

            raise PowerPointAutomationError("fake open failure")
        self.opened.append(str(path))

    def close(self) -> None:
        self.closed += 1

    def quit(self) -> None:
        self.quit_calls += 1

    def cleanup_orphans(self) -> int:
        return 0
```

Add `"FakePowerPointService"` to `__all__`.

- [ ] **Step 7: Run the tests to verify they pass**

Run: `pytest tests/test_process_utils.py tests/test_windows_com.py tests/test_powerpoint_service.py -v`
Expected: PASS

- [ ] **Step 8: Commit**

```bash
git add app/utils/process_utils.py app/integrations/windows_com.py app/services/powerpoint_service.py tests/test_process_utils.py tests/test_windows_com.py tests/test_powerpoint_service.py tests/fakes.py
git commit -m "feat: isolate PowerPoint COM behind a service and fakes"
```

---

### Task 4: iSpring publisher adapter and probe script

**Files:**
- Create: `app/services/ispring_service.py`
- Create: `scripts/probe_ispring.py`
- Create: `tests/fixtures/ispring_output/index.html`
- Create: `tests/fixtures/ispring_output/data/app.js`
- Test: `tests/test_ispring_service.py`
- Modify: `tests/fakes.py` (add `FakeISpringPublisher`)

**Interfaces:**
- Consumes: `Settings.ispring_adapter`, `Settings.ispring_install_path`, `ISpringPublishingError`, `ISpringNotConfiguredError`, `ISpringTimeoutError`
- Produces: `ISpringAvailability(installed: bool, detail: str)`, `ISpringPublisher` protocol with `is_available() -> ISpringAvailability` and `publish(pptx: Path, output_dir: Path, timeout_s: int) -> Path`, `get_publisher(settings: Settings) -> ISpringPublisher`

Do not invent a ProgID, executable name, or ribbon control. The `vba`, `uia`, and `cli` adapters raise `ISpringNotConfiguredError` describing the facts still needed.

- [ ] **Step 1: Create the fixture package**

```bash
mkdir -p tests/fixtures/ispring_output/data
```

`tests/fixtures/ispring_output/index.html`:

```html
<!DOCTYPE html>
<html><head><title>fixture</title></head>
<body><script src="data/app.js"></script></body>
</html>
```

`tests/fixtures/ispring_output/data/app.js`:

```javascript
window.ISPRING_FIXTURE = true;
```

- [ ] **Step 2: Write the failing test `tests/test_ispring_service.py`**

```python
from pathlib import Path

import pytest

from app.core.config import Settings
from app.exceptions import ISpringNotConfiguredError
from app.services.ispring_service import get_publisher
from app.utils.output_validation import validate_ispring_output


def _settings(tmp_path: Path, adapter: str) -> Settings:
    return Settings(
        app_env="development",
        worker_api_key="test-key",
        temp_root=str(tmp_path / "jobs"),
        log_root=str(tmp_path / "logs"),
        ispring_adapter=adapter,
    )


def test_not_configured_is_the_default_and_raises(tmp_path):
    publisher = get_publisher(_settings(tmp_path, "not_configured"))
    availability = publisher.is_available()
    assert availability.installed is False
    with pytest.raises(ISpringNotConfiguredError) as exc:
        publisher.publish(tmp_path / "source.pptx", tmp_path / "out", timeout_s=1)
    assert "probe_ispring.py" in str(exc.value)


@pytest.mark.parametrize("adapter", ["vba", "uia", "cli"])
def test_unverified_adapters_do_not_invent_an_api(tmp_path, adapter):
    publisher = get_publisher(_settings(tmp_path, adapter))
    with pytest.raises(ISpringNotConfiguredError):
        publisher.publish(tmp_path / "source.pptx", tmp_path / "out", timeout_s=1)


def test_fake_copies_the_fixture_package(tmp_path):
    output = tmp_path / "out"
    publisher = get_publisher(_settings(tmp_path, "fake"))
    result = publisher.publish(tmp_path / "source.pptx", output, timeout_s=1)
    assert result == output
    validate_ispring_output(output)
    assert publisher.is_available().installed is True
```

- [ ] **Step 3: Run the test to verify it fails**

Run: `pytest tests/test_ispring_service.py -v`
Expected: FAIL with `ModuleNotFoundError`

- [ ] **Step 4: Write `app/services/ispring_service.py`**

```python
"""iSpring publishing adapters.

The real publishing mechanism is unknown until scripts/probe_ispring.py is
run on the installed Cloud PC version. This module therefore ships a
loud default and three empty adapters. Do not add guessed ProgIDs,
executables, or ribbon identifiers here.
"""

from __future__ import annotations

import logging
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Protocol

from app.core.config import Settings
from app.exceptions import ISpringNotConfiguredError
from app.schemas.health import CheckResult

logger = logging.getLogger(__name__)

FIXTURE_DIR = Path(__file__).resolve().parents[2] / "tests" / "fixtures" / "ispring_output"

_PROBE_HINT = (
    "iSpring publishing is not configured. Run scripts/probe_ispring.py on the "
    "Windows Cloud PC, then set ISPRING_ADAPTER to the adapter named in that "
    "report and implement only the verified interface."
)


@dataclass(frozen=True)
class ISpringAvailability:
    installed: bool
    detail: str

    def as_check(self) -> CheckResult:
        return CheckResult(name="ispring", ok=self.installed, detail=self.detail)


class ISpringPublisher(Protocol):
    def is_available(self) -> ISpringAvailability: ...

    def publish(self, pptx: Path, output_dir: Path, timeout_s: int) -> Path: ...


class NotConfiguredPublisher:
    def is_available(self) -> ISpringAvailability:
        return ISpringAvailability(installed=False, detail=_PROBE_HINT)

    def publish(self, pptx: Path, output_dir: Path, timeout_s: int) -> Path:
        raise ISpringNotConfiguredError(_PROBE_HINT)


class UnverifiedAdapter(NotConfiguredPublisher):
    """vba / uia / cli placeholders. Same failure, more specific naming."""

    def __init__(self, name: str) -> None:
        self._name = name

    def is_available(self) -> ISpringAvailability:
        return ISpringAvailability(
            installed=False,
            detail=f"ISPRING_ADAPTER={self._name} has no verified implementation. {_PROBE_HINT}",
        )

    def publish(self, pptx: Path, output_dir: Path, timeout_s: int) -> Path:
        raise ISpringNotConfiguredError(
            f"ISPRING_ADAPTER={self._name} is a stub. {_PROBE_HINT}"
        )


class FakePublisher:
    """Copies the checked-in fixture package. Used by unit tests only."""

    def is_available(self) -> ISpringAvailability:
        return ISpringAvailability(installed=True, detail="fake publisher")

    def publish(self, pptx: Path, output_dir: Path, timeout_s: int) -> Path:
        if not FIXTURE_DIR.exists():
            raise ISpringNotConfiguredError(
                f"fake publisher fixture is missing: {FIXTURE_DIR}"
            )
        if output_dir.exists():
            shutil.rmtree(output_dir)
        shutil.copytree(FIXTURE_DIR, output_dir)
        logger.info("fake iSpring publish copied fixture", extra={"stage": "ispring_publish"})
        return output_dir


def get_publisher(settings: Settings) -> ISpringPublisher:
    choice = settings.ispring_adapter
    if choice == "fake":
        return FakePublisher()
    if choice == "not_configured":
        return NotConfiguredPublisher()
    if choice in {"vba", "uia", "cli"}:
        return UnverifiedAdapter(choice)
    raise ISpringNotConfiguredError(f"unknown ISPRING_ADAPTER {choice!r}. {_PROBE_HINT}")
```

- [ ] **Step 5: Write `scripts/probe_ispring.py`**

This script gathers facts. It must not call any unpublished iSpring API. It writes a report under `LOG_ROOT` or `./logs`.

```python
"""Gather facts about the installed iSpring / PowerPoint add-in.

Run this on the Windows Cloud PC, then attach the report before implementing
a real ISPRING_ADAPTER. This script never publishes a presentation.
"""

from __future__ import annotations

import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path


def _safe(fn, default=None):
    try:
        return fn()
    except Exception as exc:  # noqa: BLE001 — probe must keep going
        return {"error": f"{type(exc).__name__}: {exc}"}


def probe_com_addins() -> list[dict]:
    import win32com.client  # type: ignore

    app = win32com.client.Dispatch("PowerPoint.Application")
    rows = []
    for addin in app.COMAddIns:
        rows.append(
            {
                "prog_id": str(getattr(addin, "ProgId", "")),
                "description": str(getattr(addin, "Description", "")),
                "connect": bool(getattr(addin, "Connect", False)),
                "guid": str(getattr(addin, "Guid", "")),
            }
        )
    return rows


def probe_install_paths() -> list[dict]:
    candidates = []
    configured = os.environ.get("ISPRING_INSTALL_PATH", "")
    roots = [
        configured,
        r"C:\Program Files\iSpring",
        r"C:\Program Files (x86)\iSpring",
    ]
    for root in roots:
        if not root:
            continue
        path = Path(root)
        if not path.exists():
            candidates.append({"path": root, "exists": False})
            continue
        executables = [
            {"path": str(item), "name": item.name}
            for item in path.rglob("*")
            if item.is_file() and item.suffix.lower() in {".exe", ".dll"}
        ]
        candidates.append(
            {"path": root, "exists": True, "binaries_sample": executables[:50]}
        )
    return candidates


def probe_registry() -> list[str]:
    try:
        import winreg  # type: ignore
    except ImportError:
        return ["winreg unavailable"]
    keys = [
        r"SOFTWARE\iSpring",
        r"SOFTWARE\WOW6432Node\iSpring",
    ]
    found = []
    for hive, hive_name in ((winreg.HKEY_LOCAL_MACHINE, "HKLM"), (winreg.HKEY_CURRENT_USER, "HKCU")):
        for key in keys:
            try:
                handle = winreg.OpenKey(hive, key)
            except OSError:
                continue
            found.append(f"{hive_name}\\{key}")
            handle.Close()
    return found


def probe_application_run() -> dict:
    return {
        "question": "Does Application.Run expose any iSpring macros?",
        "action": "Inspect the COMAddIns report and try Application.Run only for names listed there.",
        "do_not": "Do not guess macro names.",
    }


def main() -> int:
    if sys.platform != "win32":
        print("This probe needs Windows with PowerPoint and iSpring installed.", file=sys.stderr)
        return 2

    report = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "com_addins": _safe(probe_com_addins, []),
        "install_paths": _safe(probe_install_paths, []),
        "registry_keys": _safe(probe_registry, []),
        "application_run": probe_application_run(),
        "needed_to_finish_integration": [
            "Which COMAddIn ProgID is iSpring, and is Connect true?",
            "Is there a callable Application.Run entry or only ribbon UI?",
            "Where does a manual publish write index.html?",
            "How long does a typical publish take (sets ISPRING_PUBLISH_TIMEOUT_SECONDS)?",
        ],
    }
    log_root = Path(os.environ.get("LOG_ROOT", "logs"))
    log_root.mkdir(parents=True, exist_ok=True)
    out = log_root / "ispring-probe.json"
    out.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(f"Wrote {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 6: Append `FakeISpringPublisher` to `tests/fakes.py`**

```python
@dataclass
class FakeISpringPublisher:
    available: bool = True
    published: list[str] = field(default_factory=list)
    fail: bool = False
    timeout: bool = False

    def is_available(self):
        from app.services.ispring_service import ISpringAvailability

        return ISpringAvailability(installed=self.available, detail="fake")

    def publish(self, pptx, output_dir, timeout_s: int):
        from pathlib import Path

        from app.exceptions import ISpringPublishingError, ISpringTimeoutError

        if self.timeout:
            raise ISpringTimeoutError("fake timeout")
        if self.fail:
            raise ISpringPublishingError("fake publish failure")
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        (output_dir / "data").mkdir(exist_ok=True)
        (output_dir / "index.html").write_text("<html>ok</html>", encoding="utf-8")
        (output_dir / "data" / "app.js").write_text("ok", encoding="utf-8")
        self.published.append(str(pptx))
        return output_dir
```

Add `"FakeISpringPublisher"` to `__all__`.

- [ ] **Step 7: Run the test to verify it passes**

Run: `pytest tests/test_ispring_service.py -v`
Expected: PASS

- [ ] **Step 8: Commit**

```bash
git add app/services/ispring_service.py scripts/probe_ispring.py tests/fixtures/ispring_output tests/test_ispring_service.py tests/fakes.py
git commit -m "feat: add iSpring publisher adapters and a Windows probe script"
```

---

### Task 5: Storage backends

**Files:**
- Create: `app/services/storage_service.py`
- Create: `app/integrations/storage_backends/__init__.py`
- Create: `app/integrations/storage_backends/local_fs.py`
- Create: `app/integrations/storage_backends/s3.py`
- Create: `app/integrations/storage_backends/http_api.py`
- Test: `tests/test_storage_service.py`

**Interfaces:**
- Consumes: `Settings.storage_backend` and the `S3_*` / `LOCAL_STORAGE_*` fields, `iter_files`, `OutputUploadError`, `create_sync_client`
- Produces: `UploadResult(output_path: str, entry_file: str, iframe_url: str)`, `StorageBackend` protocol with `upload_directory(local: Path, prefix: str) -> UploadResult` and `verify(entry_url: str) -> bool`, `mime_type_for(path: Path) -> str`, `get_storage_backend(settings: Settings) -> StorageBackend`

`entry_file` is always `"index.html"`. `iframe_url` is the absolute URL a browser would load. Input download is NOT part of this interface.

- [ ] **Step 1: Write the failing test `tests/test_storage_service.py`**

```python
from pathlib import Path

import httpx
import pytest

from app.core.config import Settings
from app.exceptions import OutputUploadError
from app.services.storage_service import get_storage_backend, mime_type_for
from app.utils.output_validation import validate_ispring_output


def _package(root: Path) -> Path:
    (root / "data").mkdir(parents=True)
    (root / "index.html").write_text("<html>ok</html>", encoding="utf-8")
    (root / "data" / "app.js").write_text("console.log(1)", encoding="utf-8")
    return root


def _settings(tmp_path: Path, **overrides) -> Settings:
    values = {
        "app_env": "development",
        "worker_api_key": "test-key",
        "temp_root": str(tmp_path / "jobs"),
        "log_root": str(tmp_path / "logs"),
        "storage_backend": "local_fs",
        "local_storage_root": str(tmp_path / "cdn"),
        "local_storage_public_base_url": "https://cdn.example.com/ppt",
    }
    values.update(overrides)
    return Settings(**values)


def test_mime_types_are_explicit():
    assert mime_type_for(Path("x.js")) == "application/javascript"
    assert mime_type_for(Path("x.wasm")) == "application/wasm"
    assert mime_type_for(Path("x.html")) == "text/html"


def test_local_fs_uploads_tree_and_builds_iframe_url(tmp_path):
    local = _package(tmp_path / "output")
    backend = get_storage_backend(_settings(tmp_path))
    result = backend.upload_directory(local, prefix="materials/5001/101")
    assert result.entry_file == "index.html"
    assert result.output_path == "materials/5001/101"
    assert result.iframe_url == "https://cdn.example.com/ppt/materials/5001/101/index.html"
    copied = Path(tmp_path / "cdn" / "materials" / "5001" / "101")
    validate_ispring_output(copied)


def test_local_fs_verify_uses_injected_client(tmp_path):
    def handler(request: httpx.Request) -> httpx.Response:
        assert request.method == "HEAD"
        return httpx.Response(200)

    backend = get_storage_backend(_settings(tmp_path))
    assert backend.verify(
        "https://cdn.example.com/ppt/materials/5001/101/index.html",
        client=httpx.Client(transport=httpx.MockTransport(handler)),
    )


def test_http_api_backend_is_a_documented_stub(tmp_path):
    settings = _settings(tmp_path, storage_backend="http_api")
    backend = get_storage_backend(settings)
    with pytest.raises(OutputUploadError) as exc:
        backend.upload_directory(_package(tmp_path / "output"), prefix="x")
    assert "not implemented" in str(exc.value).lower()
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pytest tests/test_storage_service.py -v`
Expected: FAIL with `ModuleNotFoundError`

- [ ] **Step 3: Write `app/services/storage_service.py`**

```python
"""Storage factory and shared upload types."""

from __future__ import annotations

import mimetypes
from dataclasses import dataclass
from pathlib import Path
from typing import Protocol

from app.core.config import Settings
from app.exceptions import OutputUploadError

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


@dataclass(frozen=True)
class UploadResult:
    output_path: str
    entry_file: str
    iframe_url: str


class StorageBackend(Protocol):
    def upload_directory(self, local: Path, prefix: str) -> UploadResult: ...

    def verify(self, entry_url: str, *, client=None) -> bool: ...


def mime_type_for(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in MIME_BY_SUFFIX:
        return MIME_BY_SUFFIX[suffix]
    guessed, _ = mimetypes.guess_type(path.name)
    return guessed or "application/octet-stream"


def get_storage_backend(settings: Settings) -> StorageBackend:
    if settings.storage_backend == "local_fs":
        from app.integrations.storage_backends.local_fs import LocalFsStorageBackend

        return LocalFsStorageBackend(settings)
    if settings.storage_backend == "s3":
        from app.integrations.storage_backends.s3 import S3StorageBackend

        return S3StorageBackend(settings)
    if settings.storage_backend == "http_api":
        from app.integrations.storage_backends.http_api import HttpApiStorageBackend

        return HttpApiStorageBackend(settings)
    raise OutputUploadError(
        f"unknown STORAGE_BACKEND {settings.storage_backend!r}",
        user_message="Storage is not configured.",
    )
```

- [ ] **Step 4: Write the three backends**

`app/integrations/storage_backends/__init__.py` — empty.

`app/integrations/storage_backends/local_fs.py`:

```python
from __future__ import annotations

import logging
import shutil
from pathlib import Path

import httpx

from app.core.config import Settings
from app.core.logging_config import scrub_url
from app.exceptions import OutputUploadError
from app.integrations.http_client import create_sync_client
from app.services.storage_service import UploadResult

logger = logging.getLogger(__name__)


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
        logger.info("local_fs upload finished", extra={"stage": "upload", "prefix": prefix})
        return UploadResult(
            output_path=prefix.strip("/"),
            entry_file="index.html",
            iframe_url=iframe_url,
        )

    def verify(self, entry_url: str, *, client: httpx.Client | None = None) -> bool:
        return _head_or_range(entry_url, self._settings.upload_timeout_seconds, client)
```

`app/integrations/storage_backends/s3.py`:

```python
from __future__ import annotations

import logging
from pathlib import Path

import httpx

from app.core.config import Settings
from app.core.logging_config import scrub_url
from app.exceptions import OutputUploadError
from app.services.storage_service import UploadResult, mime_type_for
from app.utils.file_utils import iter_files

logger = logging.getLogger(__name__)


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
        key_prefix = "/".join(
            part for part in (self._settings.s3_key_prefix.strip("/"), prefix.strip("/")) if part
        )
        client = self._s3()
        try:
            for absolute, relative in iter_files(local):
                key = f"{key_prefix}/{relative}"
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
            f"{self._settings.s3_public_base_url.rstrip('/')}/{key_prefix}/index.html"
        )
        logger.info(
            "s3 upload finished",
            extra={"stage": "upload", "prefix": key_prefix, "url": scrub_url(iframe_url)},
        )
        return UploadResult(
            output_path=key_prefix,
            entry_file="index.html",
            iframe_url=iframe_url,
        )

    def verify(self, entry_url: str, *, client: httpx.Client | None = None) -> bool:
        from app.integrations.storage_backends.local_fs import _head_or_range

        return _head_or_range(entry_url, self._settings.upload_timeout_seconds, client)
```

Put `_head_or_range` in `local_fs.py` (already referenced) as a module-level function both backends can import:

```python
def _head_or_range(entry_url: str, timeout_s: float, client: httpx.Client | None) -> bool:
    from app.integrations.http_client import create_sync_client

    owns = client is None
    client = client or create_sync_client(timeout_s=timeout_s)
    try:
        head = client.head(entry_url)
        if head.status_code < 400:
            return True
        ranged = client.get(entry_url, headers={"Range": "bytes=0-63"})
        return ranged.status_code < 400
    except httpx.HTTPError:
        logger.warning("storage verify failed", extra={"url": scrub_url(entry_url)}, exc_info=True)
        return False
    finally:
        if owns:
            client.close()
```

`app/integrations/storage_backends/http_api.py`:

```python
from __future__ import annotations

from pathlib import Path

from app.core.config import Settings
from app.exceptions import OutputUploadError
from app.services.storage_service import UploadResult


_FACTS = (
    "STORAGE_BACKEND=http_api is not implemented. Needed facts: "
    "endpoint path, authentication scheme, whether the upload is multipart "
    "files or a zip, and the success response shape including the iframe URL."
)


class HttpApiStorageBackend:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings

    def upload_directory(self, local: Path, prefix: str) -> UploadResult:
        raise OutputUploadError(_FACTS, user_message="HTTP API storage is not implemented.")

    def verify(self, entry_url: str, *, client=None) -> bool:
        raise OutputUploadError(_FACTS, user_message="HTTP API storage is not implemented.")
```

- [ ] **Step 5: Run the test to verify it passes**

Run: `pytest tests/test_storage_service.py -v`
Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add app/services/storage_service.py app/integrations/storage_backends tests/test_storage_service.py
git commit -m "feat: add storage adapters for local_fs, s3, and http_api"
```

---

### Task 6: Backend client, Slack, and orphaned-output recovery

**Files:**
- Create: `app/services/backend_service.py`
- Create: `app/services/slack_service.py`
- Test: `tests/test_backend_service.py`
- Test: `tests/test_slack_service.py`
- Modify: `tests/fakes.py` (add `FakeBackendClient`, `FakeSlackNotifier`)

**Interfaces:**
- Consumes: `JobRow`, `UploadResult`, `Settings.backend_*`, `Settings.slack_*`, `Settings.log_root`, `BackendUpdateError`, `create_sync_client`, `scrub_url`
- Produces: `BackendClient.update_material_output(job: JobRow, result: UploadResult) -> None`, `record_orphaned_output(log_root: str, job: JobRow, result: UploadResult) -> Path`, `SlackNotifier.notify_failure(job: JobRow, *, stage: str, error_message: str) -> None` (never raises)

The backend POST body is exactly:

```json
{
  "material_id": 5001,
  "job_id": 101,
  "output_path": "materials/5001/101",
  "entry_file": "index.html",
  "iframe_url": "https://cdn.example.com/ppt/materials/5001/101/index.html"
}
```

The existing backend stores `iframe_url`. This worker's table does not.

- [ ] **Step 1: Write the failing tests**

`tests/test_backend_service.py`:

```python
import json
from pathlib import Path

import httpx
import pytest

from app.core.config import Settings
from app.exceptions import BackendUpdateError
from app.services.backend_service import BackendClient, record_orphaned_output
from app.services.storage_service import UploadResult
from tests.fakes import make_job_row


def _settings(tmp_path: Path, **overrides) -> Settings:
    values = {
        "app_env": "development",
        "worker_api_key": "test-key",
        "temp_root": str(tmp_path / "jobs"),
        "log_root": str(tmp_path / "logs"),
        "backend_base_url": "https://backend.example.com",
        "backend_update_path": "/internal/ppt-jobs/complete",
        "backend_api_key": "backend-secret",
        "backend_api_key_header": "X-API-Key",
        "backend_request_timeout_seconds": 5,
    }
    values.update(overrides)
    return Settings(**values)


RESULT = UploadResult(
    output_path="materials/5001/101",
    entry_file="index.html",
    iframe_url="https://cdn.example.com/ppt/materials/5001/101/index.html",
)


def test_posts_expected_body_and_header(tmp_path):
    seen = {}

    def handler(request: httpx.Request) -> httpx.Response:
        seen["url"] = str(request.url)
        seen["header"] = request.headers.get("x-api-key")
        seen["body"] = json.loads(request.content)
        return httpx.Response(200, json={"ok": True})

    client = BackendClient(
        _settings(tmp_path),
        client=httpx.Client(transport=httpx.MockTransport(handler)),
    )
    client.update_material_output(make_job_row(status=1), RESULT)
    assert seen["url"] == "https://backend.example.com/internal/ppt-jobs/complete"
    assert seen["header"] == "backend-secret"
    assert seen["body"]["job_id"] == 101
    assert seen["body"]["material_id"] == 5001
    assert seen["body"]["iframe_url"] == RESULT.iframe_url
    assert seen["body"]["output_path"] == RESULT.output_path
    assert seen["body"]["entry_file"] == "index.html"


def test_non_2xx_raises(tmp_path):
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(500, text="nope")

    client = BackendClient(
        _settings(tmp_path),
        client=httpx.Client(transport=httpx.MockTransport(handler)),
    )
    with pytest.raises(BackendUpdateError):
        client.update_material_output(make_job_row(status=1), RESULT)


def test_missing_config_raises(tmp_path):
    client = BackendClient(_settings(tmp_path, backend_base_url="", backend_update_path=""))
    with pytest.raises(BackendUpdateError):
        client.update_material_output(make_job_row(status=1), RESULT)


def test_orphaned_output_is_appended_outside_the_job_dir(tmp_path):
    log_root = tmp_path / "logs"
    path = record_orphaned_output(str(log_root), make_job_row(), RESULT)
    assert path == log_root / "orphaned_outputs.jsonl"
    row = json.loads(path.read_text(encoding="utf-8").splitlines()[0])
    assert row["job_id"] == 101
    assert row["iframe_url"] == RESULT.iframe_url
    assert "timestamp" in row
```

`tests/test_slack_service.py`:

```python
import json
from pathlib import Path

import httpx

from app.core.config import Settings
from app.services.slack_service import SlackNotifier
from tests.fakes import make_job_row


def _settings(tmp_path: Path, **overrides) -> Settings:
    values = {
        "app_env": "development",
        "worker_api_key": "test-key",
        "temp_root": str(tmp_path / "jobs"),
        "log_root": str(tmp_path / "logs"),
        "slack_webhook_url": "https://hooks.slack.com/services/T/B/xxx",
        "slack_request_timeout_seconds": 5,
    }
    values.update(overrides)
    return Settings(**values)


def test_posts_failure_fields_and_never_logs_webhook(tmp_path, caplog):
    seen = {}

    def handler(request: httpx.Request) -> httpx.Response:
        seen["body"] = json.loads(request.content)
        return httpx.Response(200, text="ok")

    SlackNotifier(
        _settings(tmp_path),
        client=httpx.Client(transport=httpx.MockTransport(handler)),
    ).notify_failure(
        make_job_row(),
        stage="upload",
        error_message="s3 denied",
    )
    text = seen["body"]["text"]
    assert "101" in text
    assert "5001" in text
    assert "Introduction to Biology" in text
    assert "ABC College" in text
    assert "upload" in text
    assert "s3 denied" in text
    assert "hooks.slack.com" not in caplog.text
    assert "T/B/xxx" not in caplog.text


def test_slack_failure_does_not_raise(tmp_path):
    def handler(request: httpx.Request) -> httpx.Response:
        raise httpx.ConnectError("down")

    SlackNotifier(
        _settings(tmp_path),
        client=httpx.Client(transport=httpx.MockTransport(handler)),
    ).notify_failure(make_job_row(), stage="download", error_message="x")


def test_empty_webhook_is_a_noop(tmp_path):
    SlackNotifier(_settings(tmp_path, slack_webhook_url="")).notify_failure(
        make_job_row(), stage="download", error_message="x"
    )
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `pytest tests/test_backend_service.py tests/test_slack_service.py -v`
Expected: FAIL with `ModuleNotFoundError`

- [ ] **Step 3: Write `app/services/backend_service.py`**

```python
"""Notify the existing PHP backend. This worker table does not store iframe_url."""

from __future__ import annotations

import json
import logging
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin

import httpx

from app.core.config import Settings
from app.core.logging_config import scrub_url
from app.exceptions import BackendUpdateError
from app.integrations.database import JobRow
from app.integrations.http_client import create_sync_client
from app.services.storage_service import UploadResult

logger = logging.getLogger(__name__)


class BackendClient:
    def __init__(self, settings: Settings, *, client: httpx.Client | None = None) -> None:
        self._settings = settings
        self._client = client

    def update_material_output(self, job: JobRow, result: UploadResult) -> None:
        base = self._settings.backend_base_url
        path = self._settings.backend_update_path
        if not base or not path:
            raise BackendUpdateError(
                "BACKEND_BASE_URL and BACKEND_UPDATE_PATH must both be set",
                user_message="The existing backend is not configured.",
            )
        url = urljoin(base.rstrip("/") + "/", path.lstrip("/"))
        payload = {
            "material_id": job.material_id,
            "job_id": job.id,
            "output_path": result.output_path,
            "entry_file": result.entry_file,
            "iframe_url": result.iframe_url,
        }
        headers = {"Content-Type": "application/json"}
        if self._settings.backend_api_key:
            headers[self._settings.backend_api_key_header] = self._settings.backend_api_key

        owns = self._client is None
        client = self._client or create_sync_client(
            timeout_s=self._settings.backend_request_timeout_seconds
        )
        try:
            response = client.post(url, json=payload, headers=headers)
        except httpx.HTTPError as exc:
            logger.error(
                "backend update transport failed",
                extra={"url": scrub_url(url), "stage": "backend_update"},
                exc_info=True,
            )
            raise BackendUpdateError(
                f"backend update failed: {exc}",
                user_message="The existing backend could not be updated.",
            ) from exc
        finally:
            if owns:
                client