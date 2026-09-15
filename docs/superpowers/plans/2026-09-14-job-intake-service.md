# Job Intake Service Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build the FastAPI intake half of the PPT automation worker — authenticated job submission, atomic single-claim semantics, single-slot concurrency, health reporting, and status lookup — with the processing pipeline behind an injected protocol.

**Architecture:** Single-process FastAPI app on uvicorn with exactly one worker. Imports flow `api -> services -> integrations`. The endpoint acquires an in-process `threading.Lock` slot *before* touching the database, then performs a conditional `UPDATE` whose affected-row count is the sole authority on whether the job was claimed. Work is scheduled as a tracked asyncio task and the response returns `202` immediately.

**Tech Stack:** Python 3.11+, FastAPI, uvicorn, pydantic v2, pydantic-settings, SQLAlchemy 2.x Core (not the ORM), PyMySQL, httpx, pytest.

## Global Constraints

- Design spec: `docs/superpowers/specs/2026-09-14-ppt-automation-worker-design.md`. Read it before starting.
- The entire test suite MUST run and pass on macOS and Linux with no PowerPoint, no iSpring, and no MySQL installed.
- No module in `app/api/` or `app/services/` may import a Windows-only module (`win32com`, `pythoncom`, `pywintypes`, `pywinauto`) at module scope.
- Write NO DDL. No `CREATE TABLE`, `ALTER TABLE`, `CREATE INDEX`, no migrations. The `ppt_automation_jobs` table already exists.
- The job table has exactly these columns: `id`, `material_id`, `material_name`, `institution_name`, `ppt_file_url`, `status`, `error_message`, `created_at`, `updated_at`. Never reference `current_step`, `worker_type`, `retry_count`, `retry_at`, or `iframe_url`.
- Status values: `0` PENDING, `1` PROCESSING, `2` COMPLETED, `3` FAILED.
- All SQL uses SQLAlchemy `text()` with bound parameters. Never format or concatenate a value into a SQL string.
- No bare `except:`. No `except Exception` without `exc_info=True` on the log call.
- Never log an API key, password, token, webhook URL, or a URL's query string.
- Every filesystem path is derived from the integer `job_id`. Never from `material_name`, `institution_name`, or a remote filename.
- Copy the claim SQL character for character from Task 5. Do not re-derive it.
- Commit after every task.

---

### Task 1: Project scaffolding and configuration

**Files:**
- Create: `requirements.txt`
- Create: `requirements-dev.txt`
- Create: `.gitignore`
- Create: `.env.example`
- Create: `app/__init__.py`, `app/core/__init__.py`, `app/api/__init__.py`, `app/schemas/__init__.py`, `app/services/__init__.py`, `app/integrations/__init__.py`, `app/utils/__init__.py`, `tests/__init__.py`
- Create: `app/core/config.py`
- Create: `pytest.ini`
- Test: `tests/test_config.py`

**Interfaces:**
- Consumes: nothing.
- Produces: `app.core.config.Settings` (pydantic-settings class), `app.core.config.get_settings() -> Settings` (cached), `app.core.config.reset_settings_cache() -> None` for tests.

- [ ] **Step 1: Create the package directories and empty `__init__.py` files**

```bash
mkdir -p app/api app/core app/schemas app/services app/integrations app/utils
mkdir -p scripts tests docs temp logs
touch app/__init__.py app/api/__init__.py app/core/__init__.py app/schemas/__init__.py
touch app/services/__init__.py app/integrations/__init__.py app/utils/__init__.py
touch tests/__init__.py
```

- [ ] **Step 2: Write `requirements.txt`**

The `sys_platform == "win32"` markers are what let `pip install -r requirements.txt` succeed on a Mac or Linux dev machine. Do not remove them.

```
fastapi==0.115.6
uvicorn[standard]==0.34.0
pydantic==2.10.4
pydantic-settings==2.7.0
httpx==0.28.1
SQLAlchemy==2.0.36
PyMySQL==1.1.1
boto3==1.35.90
psutil==6.1.1
python-dotenv==1.0.1
pywin32==308; sys_platform == "win32"
pywinauto==0.6.8; sys_platform == "win32"
```

- [ ] **Step 3: Write `requirements-dev.txt`**

```
-r requirements.txt
pytest==8.3.4
pytest-asyncio==0.25.0
pytest-cov==6.0.0
```

- [ ] **Step 4: Write `.gitignore`**

```
.env
__pycache__/
*.py[cod]
.venv/
venv/
.pytest_cache/
.coverage
htmlcov/
temp/*
!temp/.gitkeep
logs/*
!logs/.gitkeep
*.pptx
*.part
.DS_Store
```

- [ ] **Step 5: Write `pytest.ini`**

```ini
[pytest]
testpaths = tests
asyncio_mode = auto
markers =
    mysql: requires a real MySQL database (skipped by default)
    windows: requires Windows with PowerPoint installed (skipped by default)
```

- [ ] **Step 6: Write the failing test `tests/test_config.py`**

```python
import pytest
from pydantic import ValidationError

from app.core.config import Settings


def _base(**overrides):
    values = {
        "app_env": "development",
        "worker_api_key": "test-key",
        "database_url": "mysql+pymysql://u:p@localhost:3306/db",
        "temp_root": "/tmp/ppt-automation/jobs",
        "log_root": "/tmp/ppt-automation/logs",
    }
    values.update(overrides)
    return values


def test_defaults_are_safe():
    s = Settings(**_base())
    assert s.ispring_adapter == "not_configured"
    assert s.storage_backend == "local_fs"
    assert s.keep_failed_job_files is False
    assert s.powerpoint_kill_orphans is False
    assert s.uvicorn_workers == 1
    assert s.max_ppt_size_mb == 500
    assert s.download_allowed_hosts == []


def test_max_ppt_size_bytes_is_derived():
    s = Settings(**_base(max_ppt_size_mb=2))
    assert s.max_ppt_size_bytes == 2 * 1024 * 1024


def test_download_allowed_hosts_parses_comma_separated():
    s = Settings(**_base(download_allowed_hosts="cdn.example.com, storage.example.com"))
    assert s.download_allowed_hosts == ["cdn.example.com", "storage.example.com"]


def test_production_rejects_placeholder_api_key():
    with pytest.raises(ValidationError, match="WORKER_API_KEY"):
        Settings(**_base(app_env="production", worker_api_key="change-me"))


def test_production_rejects_empty_api_key():
    with pytest.raises(ValidationError, match="WORKER_API_KEY"):
        Settings(**_base(app_env="production", worker_api_key=""))


def test_production_rejects_local_fs_storage():
    with pytest.raises(ValidationError, match="STORAGE_BACKEND"):
        Settings(**_base(app_env="production", storage_backend="local_fs"))


def test_production_allows_valid_configuration():
    s = Settings(**_base(app_env="production", storage_backend="s3"))
    assert s.app_env == "production"


def test_uvicorn_workers_cannot_exceed_one():
    with pytest.raises(ValidationError, match="UVICORN_WORKERS"):
        Settings(**_base(uvicorn_workers=2))
```

- [ ] **Step 7: Run the test to verify it fails**

Run: `pytest tests/test_config.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'app.core.config'`

- [ ] **Step 8: Write `app/core/config.py`**

`uvicorn_workers` is validated to be exactly 1 because the single-PowerPoint guarantee is an in-process lock; a second worker process would create a second lock. This is the cheapest possible place to prevent that mistake.

```python
"""Application configuration loaded from the environment."""

from __future__ import annotations

from functools import lru_cache
from typing import Literal

from pydantic import Field, computed_field, field_validator, model_validator
from pydantic_settings import BaseSettings, SettingsConfigDict

PLACEHOLDER_API_KEY = "change-me"


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,
        extra="ignore",
    )

    # Application
    app_env: Literal["development", "staging", "production"] = "development"
    app_host: str = "0.0.0.0"
    app_port: int = 8000
    uvicorn_workers: int = 1
    log_level: str = "INFO"

    # Authentication
    worker_api_key: str = ""

    # Database
    database_url: str = ""

    # Existing backend
    backend_base_url: str = ""
    backend_api_key: str = ""
    backend_update_path: str = ""
    backend_api_key_header: str = "X-API-Key"

    # Filesystem
    temp_root: str = "D:\\ppt-automation\\jobs"
    log_root: str = "D:\\ppt-automation\\logs"
    keep_failed_job_files: bool = False
    min_free_disk_gb: int = 10

    # Limits and timeouts (seconds)
    max_ppt_size_mb: int = 500
    download_timeout_seconds: int = 300
    powerpoint_start_timeout_seconds: int = 60
    powerpoint_open_timeout_seconds: int = 120
    ispring_publish_timeout_seconds: int = 1800
    upload_timeout_seconds: int = 600
    backend_request_timeout_seconds: int = 60
    slack_request_timeout_seconds: int = 15

    # Slack
    slack_webhook_url: str = ""

    # Windows tooling
    powerpoint_exe_path: str = ""
    ispring_install_path: str = ""
    powerpoint_kill_orphans: bool = False
    powerpoint_orphan_max_age_seconds: int = 3600

    # Adapters
    ispring_adapter: Literal["not_configured", "fake", "vba", "uia", "cli"] = "not_configured"
    storage_backend: Literal["local_fs", "s3", "http_api"] = "local_fs"

    # Download safety
    download_allowed_hosts: list[str] = Field(default_factory=list)

    # S3-compatible storage
    s3_bucket: str = ""
    s3_region: str = ""
    s3_endpoint_url: str = ""
    s3_access_key_id: str = ""
    s3_secret_access_key: str = ""
    s3_key_prefix: str = ""
    s3_public_base_url: str = ""

    # Local filesystem storage (development only)
    local_storage_root: str = ""
    local_storage_public_base_url: str = ""

    @field_validator("download_allowed_hosts", mode="before")
    @classmethod
    def _split_hosts(cls, value: object) -> object:
        if isinstance(value, str):
            return [host.strip() for host in value.split(",") if host.strip()]
        return value

    @field_validator("uvicorn_workers")
    @classmethod
    def _only_one_worker(cls, value: int) -> int:
        if value != 1:
            raise ValueError(
                "UVICORN_WORKERS must be 1. The single-PowerPoint guarantee is an "
                "in-process lock and additional workers would bypass it."
            )
        return value

    @computed_field  # type: ignore[prop-decorator]
    @property
    def max_ppt_size_bytes(self) -> int:
        return self.max_ppt_size_mb * 1024 * 1024

    @property
    def is_production(self) -> bool:
        return self.app_env == "production"

    @model_validator(mode="after")
    def _validate_production(self) -> "Settings":
        if not self.is_production:
            return self
        if not self.worker_api_key or self.worker_api_key == PLACEHOLDER_API_KEY:
            raise ValueError(
                "WORKER_API_KEY must be set to a real secret when APP_ENV=production."
            )
        if self.storage_backend == "local_fs":
            raise ValueError(
                "STORAGE_BACKEND=local_fs is not permitted when APP_ENV=production: "
                "published output would be written to a directory nothing serves, "
                "producing an iframe URL that cannot resolve."
            )
        return self


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    return Settings()


def reset_settings_cache() -> None:
    """Clear the cached settings. Tests use this after changing the environment."""
    get_settings.cache_clear()
```

- [ ] **Step 9: Run the test to verify it passes**

Run: `pytest tests/test_config.py -v`
Expected: PASS, 8 tests.

- [ ] **Step 10: Write `.env.example`**

Note there is no real secret in this file and it is the only env file that is committed.

```
APP_ENV=development
APP_HOST=0.0.0.0
APP_PORT=8000
UVICORN_WORKERS=1
LOG_LEVEL=INFO
WORKER_API_KEY=change-me

DATABASE_URL=mysql+pymysql://user:password@localhost:3306/dbname

BACKEND_BASE_URL=
BACKEND_API_KEY=
BACKEND_UPDATE_PATH=
BACKEND_API_KEY_HEADER=X-API-Key

TEMP_ROOT=D:\ppt-automation\jobs
LOG_ROOT=D:\ppt-automation\logs
KEEP_FAILED_JOB_FILES=false
MIN_FREE_DISK_GB=10

MAX_PPT_SIZE_MB=500

DOWNLOAD_TIMEOUT_SECONDS=300
POWERPOINT_START_TIMEOUT_SECONDS=60
POWERPOINT_OPEN_TIMEOUT_SECONDS=120
ISPRING_PUBLISH_TIMEOUT_SECONDS=1800
UPLOAD_TIMEOUT_SECONDS=600
BACKEND_REQUEST_TIMEOUT_SECONDS=60
SLACK_REQUEST_TIMEOUT_SECONDS=15

SLACK_WEBHOOK_URL=

POWERPOINT_EXE_PATH=
ISPRING_INSTALL_PATH=
POWERPOINT_KILL_ORPHANS=false
POWERPOINT_ORPHAN_MAX_AGE_SECONDS=3600

ISPRING_ADAPTER=not_configured
STORAGE_BACKEND=local_fs

DOWNLOAD_ALLOWED_HOSTS=

S3_BUCKET=
S3_REGION=
S3_ENDPOINT_URL=
S3_ACCESS_KEY_ID=
S3_SECRET_ACCESS_KEY=
S3_KEY_PREFIX=
S3_PUBLIC_BASE_URL=

LOCAL_STORAGE_ROOT=
LOCAL_STORAGE_PUBLIC_BASE_URL=
```

- [ ] **Step 11: Add `.gitkeep` files so the empty runtime directories exist**

```bash
touch temp/.gitkeep logs/.gitkeep
```

- [ ] **Step 12: Commit**

```bash
git add requirements.txt requirements-dev.txt .gitignore .env.example pytest.ini app tests temp/.gitkeep logs/.gitkeep
git commit -m "feat: project scaffolding and configuration with production guards"
```

---

### Task 2: Exception hierarchy

**Files:**
- Create: `app/exceptions.py`
- Test: `tests/test_exceptions.py`

**Interfaces:**
- Consumes: nothing.
- Produces: `app.exceptions.PptAutomationError` with class attribute `stage: str` and instance attributes `message: str`, `user_message: str`; subclasses `JobClaimError`, `DownloadError`, `PowerPointAutomationError`, `ISpringPublishingError`, `ISpringNotConfiguredError`, `ISpringTimeoutError`, `OutputValidationError`, `OutputUploadError`, `BackendUpdateError`, `DatabaseError`, `CleanupError`, `LowDiskSpaceError`.

The single shared `stage` attribute is the point of this task. Because every exception carries the stage it came from, the failure handler needs no per-type branching to know what to write into `error_message`, what to put in the Slack message, and what to set as the log's `stage` field.

- [ ] **Step 1: Write the failing test `tests/test_exceptions.py`**

```python
import pytest

from app.exceptions import (
    BackendUpdateError,
    CleanupError,
    DatabaseError,
    DownloadError,
    ISpringNotConfiguredError,
    ISpringPublishingError,
    ISpringTimeoutError,
    JobClaimError,
    LowDiskSpaceError,
    OutputUploadError,
    OutputValidationError,
    PowerPointAutomationError,
    PptAutomationError,
)


def test_every_error_is_a_pptautomationerror():
    for cls in (
        JobClaimError,
        DownloadError,
        PowerPointAutomationError,
        ISpringPublishingError,
        ISpringNotConfiguredError,
        ISpringTimeoutError,
        OutputValidationError,
        OutputUploadError,
        BackendUpdateError,
        DatabaseError,
        CleanupError,
        LowDiskSpaceError,
    ):
        assert issubclass(cls, PptAutomationError)


def test_each_subclass_declares_a_distinct_default_stage():
    stages = {
        JobClaimError: "claim",
        DownloadError: "download",
        PowerPointAutomationError: "powerpoint",
        ISpringPublishingError: "ispring_publish",
        ISpringNotConfiguredError: "ispring_publish",
        ISpringTimeoutError: "ispring_publish",
        OutputValidationError: "output_validate",
        OutputUploadError: "upload",
        BackendUpdateError: "backend_update",
        DatabaseError: "database",
        CleanupError: "cleanup",
        LowDiskSpaceError: "prepare",
    }
    for cls, stage in stages.items():
        assert cls("boom").stage == stage


def test_timeout_is_a_publishing_error_so_handlers_can_treat_it_either_way():
    err = ISpringTimeoutError("timed out after 1800s")
    assert isinstance(err, ISpringPublishingError)


def test_message_and_str():
    err = DownloadError("connection reset")
    assert err.message == "connection reset"
    assert str(err) == "connection reset"


def test_user_message_defaults_to_message_and_can_be_overridden():
    assert DownloadError("raw detail").user_message == "raw detail"
    err = DownloadError("raw detail", user_message="Could not download the file.")
    assert err.user_message == "Could not download the file."


def test_stage_can_be_overridden_per_instance():
    err = DatabaseError("deadlock", stage="complete")
    assert err.stage == "complete"


def test_base_error_has_a_fallback_stage():
    assert PptAutomationError("boom").stage == "unexpected"
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pytest tests/test_exceptions.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'app.exceptions'`

- [ ] **Step 3: Write `app/exceptions.py`**

```python
"""Exception hierarchy for the PPT automation worker.

Every failure carries the pipeline stage it occurred in. The failure handler
reads ``stage`` rather than branching on exception type, so adding a new
exception requires no changes to the handler.
"""

from __future__ import annotations


class PptAutomationError(Exception):
    """Base class for every expected failure in this service."""

    stage: str = "unexpected"

    def __init__(
        self,
        message: str,
        *,
        stage: str | None = None,
        user_message: str | None = None,
    ) -> None:
        super().__init__(message)
        self.message = message
        if stage is not None:
            self.stage = stage
        self.user_message = user_message or message


class JobClaimError(PptAutomationError):
    stage = "claim"


class LowDiskSpaceError(PptAutomationError):
    stage = "prepare"


class DownloadError(PptAutomationError):
    stage = "download"


class PowerPointAutomationError(PptAutomationError):
    stage = "powerpoint"


class ISpringPublishingError(PptAutomationError):
    stage = "ispring_publish"


class ISpringNotConfiguredError(ISpringPublishingError):
    """Raised by the default adapter: the real mechanism is not yet verified."""


class ISpringTimeoutError(ISpringPublishingError):
    """Publishing exceeded ISPRING_PUBLISH_TIMEOUT_SECONDS."""


class OutputValidationError(PptAutomationError):
    stage = "output_validate"


class OutputUploadError(PptAutomationError):
    stage = "upload"


class BackendUpdateError(PptAutomationError):
    stage = "backend_update"


class DatabaseError(PptAutomationError):
    stage = "database"


class CleanupError(PptAutomationError):
    stage = "cleanup"
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `pytest tests/test_exceptions.py -v`
Expected: PASS, 7 tests.

- [ ] **Step 5: Commit**

```bash
git add app/exceptions.py tests/test_exceptions.py
git commit -m "feat: exception hierarchy carrying pipeline stage"
```

---

### Task 3: Structured logging with job context and redaction

**Files:**
- Create: `app/core/logging_config.py`
- Test: `tests/test_logging_config.py`

**Interfaces:**
- Consumes: `app.core.config.Settings`.
- Produces: `configure_logging(settings: Settings) -> None`; `bind_job_context(job_id: int, material_id: int) -> None`; `set_stage(stage: str) -> None`; `clear_job_context() -> None`; `scrub_url(url: str) -> str`; `JsonFormatter`; `job_log_handler(log_dir: Path) -> logging.Handler`.

`scrub_url` is the load-bearing function here. Input PPT URLs are frequently pre-signed, meaning the query string contains a credential. Logging such a URL verbatim writes a working access token to disk.

- [ ] **Step 1: Write the failing test `tests/test_logging_config.py`**

```python
import json
import logging

from app.core.config import Settings
from app.core.logging_config import (
    JsonFormatter,
    bind_job_context,
    clear_job_context,
    configure_logging,
    scrub_url,
    set_stage,
)


def _settings(**overrides):
    values = {"worker_api_key": "test-key", "log_root": "/tmp/ppt-automation/logs"}
    values.update(overrides)
    return Settings(**values)


def test_scrub_url_removes_the_query_string():
    signed = "https://s3.example.com/in/a.pptx?X-Amz-Signature=deadbeef&X-Amz-Expires=900"
    assert scrub_url(signed) == "https://s3.example.com/in/a.pptx"


def test_scrub_url_removes_userinfo_credentials():
    assert scrub_url("https://user:secret@example.com/a") == "https://example.com/a"


def test_scrub_url_leaves_a_clean_url_alone():
    assert scrub_url("https://cdn.example.com/ppt/1/index.html") == (
        "https://cdn.example.com/ppt/1/index.html"
    )


def test_scrub_url_handles_a_non_url_without_raising():
    assert scrub_url("not a url") == "not a url"


def test_formatter_emits_json_with_the_expected_keys():
    record = logging.LogRecord(
        name="app.test", level=logging.INFO, pathname=__file__, lineno=1,
        msg="hello", args=(), exc_info=None,
    )
    payload = json.loads(JsonFormatter().format(record))
    assert payload["message"] == "hello"
    assert payload["level"] == "INFO"
    assert payload["logger"] == "app.test"
    assert "timestamp" in payload


def test_job_context_is_attached_to_records(caplog):
    clear_job_context()
    bind_job_context(job_id=101, material_id=5001)
    set_stage("download")
    try:
        formatter = JsonFormatter()
        record = logging.LogRecord(
            name="app.test", level=logging.INFO, pathname=__file__, lineno=1,
            msg="downloading", args=(), exc_info=None,
        )
        from app.core.logging_config import JobContextFilter

        JobContextFilter().filter(record)
        payload = json.loads(formatter.format(record))
        assert payload["job_id"] == 101
        assert payload["material_id"] == 5001
        assert payload["stage"] == "download"
    finally:
        clear_job_context()


def test_context_is_absent_when_not_bound():
    clear_job_context()
    from app.core.logging_config import JobContextFilter

    record = logging.LogRecord(
        name="app.test", level=logging.INFO, pathname=__file__, lineno=1,
        msg="startup", args=(), exc_info=None,
    )
    JobContextFilter().filter(record)
    payload = json.loads(JsonFormatter().format(record))
    assert "job_id" not in payload


def test_extra_fields_are_included():
    record = logging.LogRecord(
        name="app.test", level=logging.INFO, pathname=__file__, lineno=1,
        msg="done", args=(), exc_info=None,
    )
    record.duration_ms = 1234
    payload = json.loads(JsonFormatter().format(record))
    assert payload["duration_ms"] == 1234


def test_configure_logging_is_idempotent(tmp_path):
    settings = _settings(log_root=str(tmp_path))
    configure_logging(settings)
    count = len(logging.getLogger().handlers)
    configure_logging(settings)
    assert len(logging.getLogger().handlers) == count
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pytest tests/test_logging_config.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'app.core.logging_config'`

- [ ] **Step 3: Write `app/core/logging_config.py`**

```python
"""Structured JSON logging with per-job context and URL redaction."""

from __future__ import annotations

import json
import logging
import logging.handlers
from contextvars import ContextVar
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlsplit, urlunsplit

from app.core.config import Settings

_job_id: ContextVar[int | None] = ContextVar("job_id", default=None)
_material_id: ContextVar[int | None] = ContextVar("material_id", default=None)
_stage: ContextVar[str | None] = ContextVar("stage", default=None)

_RESERVED = frozenset(
    logging.LogRecord("", 0, "", 0, "", (), None).__dict__.keys()
) | {"message", "asctime", "taskName"}


def bind_job_context(job_id: int, material_id: int) -> None:
    _job_id.set(job_id)
    _material_id.set(material_id)


def set_stage(stage: str) -> None:
    _stage.set(stage)


def clear_job_context() -> None:
    _job_id.set(None)
    _material_id.set(None)
    _stage.set(None)


def scrub_url(url: str) -> str:
    """Return ``url`` without its query string, fragment, or userinfo.

    Pre-signed URLs carry credentials in the query string; logging one
    verbatim writes a working access token to disk.
    """
    try:
        parts = urlsplit(url)
        if not parts.scheme or not parts.netloc:
            return url
        host = parts.hostname or ""
        if parts.port:
            host = f"{host}:{parts.port}"
        return urlunsplit((parts.scheme, host, parts.path, "", ""))
    except ValueError:
        return "<unparseable-url>"


class JobContextFilter(logging.Filter):
    """Attach the current job's context variables to each record."""

    def filter(self, record: logging.LogRecord) -> bool:
        for attr, var in (
            ("job_id", _job_id),
            ("material_id", _material_id),
            ("stage", _stage),
        ):
            value = var.get()
            if value is not None:
                setattr(record, attr, value)
        return True


class JsonFormatter(logging.Formatter):
    def format(self, record: logging.LogRecord) -> str:
        payload: dict[str, object] = {
            "timestamp": datetime.fromtimestamp(
                record.created, tz=timezone.utc
            ).isoformat(),
            "level": record.levelname,
            "logger": record.name,
            "message": record.getMessage(),
        }
        for key, value in record.__dict__.items():
            if key not in _RESERVED and not key.startswith("_"):
                payload[key] = value
        if record.exc_info:
            payload["exception"] = self.formatException(record.exc_info)
        return json.dumps(payload, default=str)


def job_log_handler(log_dir: Path) -> logging.Handler:
    """A handler writing this job's log into its own directory."""
    log_dir.mkdir(parents=True, exist_ok=True)
    handler = logging.FileHandler(log_dir / "job.log", encoding="utf-8")
    handler.setFormatter(JsonFormatter())
    handler.addFilter(JobContextFilter())
    return handler


def configure_logging(settings: Settings) -> None:
    """Install the console and rotating-file handlers. Safe to call twice."""
    root = logging.getLogger()
    if getattr(root, "_ppt_automation_configured", False):
        root.setLevel(settings.log_level.upper())
        return

    root.setLevel(settings.log_level.upper())
    formatter = JsonFormatter()
    context_filter = JobContextFilter()

    console = logging.StreamHandler()
    console.setFormatter(formatter)
    console.addFilter(context_filter)
    root.addHandler(console)

    log_root = Path(settings.log_root)
    try:
        log_root.mkdir(parents=True, exist_ok=True)
        file_handler = logging.handlers.RotatingFileHandler(
            log_root / "worker.log",
            maxBytes=10 * 1024 * 1024,
            backupCount=10,
            encoding="utf-8",
        )
        file_handler.setFormatter(formatter)
        file_handler.addFilter(context_filter)
        root.addHandler(file_handler)
    except OSError:
        logging.getLogger(__name__).warning(
            "could not open log file, continuing with console logging only",
            extra={"log_root": str(log_root)},
            exc_info=True,
        )

    root._ppt_automation_configured = True  # type: ignore[attr-defined]
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `pytest tests/test_logging_config.py -v`
Expected: PASS, 9 tests.

- [ ] **Step 5: Commit**

```bash
git add app/core/logging_config.py tests/test_logging_config.py
git commit -m "feat: structured JSON logging with job context and URL scrubbing"
```

---

### Task 4: Path safety and validation utilities

**Files:**
- Create: `app/utils/file_utils.py`
- Create: `app/utils/validators.py`
- Test: `tests/test_file_utils.py`
- Test: `tests/test_validators.py`

**Interfaces:**
- Consumes: `app.exceptions`, `app.core.config.Settings`.
- Produces:
  - `app.utils.file_utils.JobPaths` — dataclass with fields `root: Path`, `input_dir: Path`, `working_dir: Path`, `output_dir: Path`, `log_dir: Path`, and property `source_pptx: Path`.
  - `app.utils.file_utils.build_job_paths(temp_root: str | Path, job_id: int) -> JobPaths`
  - `app.utils.file_utils.create_job_dirs(paths: JobPaths) -> None`
  - `app.utils.file_utils.assert_within(base: Path, candidate: Path) -> Path`
  - `app.utils.file_utils.free_disk_gb(path: str | Path) -> float`
  - `app.utils.file_utils.iter_files(root: Path) -> Iterator[tuple[Path, str]]` yielding `(absolute_path, posix_relative_path)`
  - `app.utils.validators.validate_download_url(url: str, allowed_hosts: list[str]) -> str`
  - `app.utils.validators.assert_looks_like_pptx(path: Path) -> None`

`build_job_paths` takes an `int`. That type signature is the path-traversal defence: a caller cannot pass `../../etc` where an integer is required, so traversal is structurally impossible rather than filtered out.

- [ ] **Step 1: Write the failing test `tests/test_file_utils.py`**

```python
import pytest

from app.utils.file_utils import (
    JobPaths,
    assert_within,
    build_job_paths,
    create_job_dirs,
    free_disk_gb,
    iter_files,
)


def test_build_job_paths_uses_the_job_id_only(tmp_path):
    paths = build_job_paths(tmp_path, 101)
    assert paths.root == tmp_path / "101"
    assert paths.input_dir == tmp_path / "101" / "input"
    assert paths.working_dir == tmp_path / "101" / "working"
    assert paths.output_dir == tmp_path / "101" / "output"
    assert paths.log_dir == tmp_path / "101" / "logs"


def test_source_pptx_has_a_fixed_name(tmp_path):
    paths = build_job_paths(tmp_path, 101)
    assert paths.source_pptx == paths.input_dir / "source.pptx"


def test_build_job_paths_rejects_a_non_integer_job_id(tmp_path):
    with pytest.raises(TypeError):
        build_job_paths(tmp_path, "../../etc")


def test_build_job_paths_rejects_a_non_positive_job_id(tmp_path):
    with pytest.raises(ValueError):
        build_job_paths(tmp_path, 0)


def test_create_job_dirs_creates_every_subdirectory(tmp_path):
    paths = build_job_paths(tmp_path, 7)
    create_job_dirs(paths)
    for directory in (paths.input_dir, paths.working_dir, paths.output_dir, paths.log_dir):
        assert directory.is_dir()


def test_create_job_dirs_is_idempotent(tmp_path):
    paths = build_job_paths(tmp_path, 7)
    create_job_dirs(paths)
    create_job_dirs(paths)
    assert paths.input_dir.is_dir()


def test_assert_within_returns_a_contained_path(tmp_path):
    inner = tmp_path / "a" / "b.txt"
    assert assert_within(tmp_path, inner) == inner.resolve()


def test_assert_within_rejects_an_escaping_path(tmp_path):
    with pytest.raises(ValueError, match="outside"):
        assert_within(tmp_path, tmp_path / ".." / "escaped.txt")


def test_free_disk_gb_is_positive(tmp_path):
    assert free_disk_gb(tmp_path) > 0


def test_iter_files_yields_posix_relative_paths(tmp_path):
    (tmp_path / "res" / "deep").mkdir(parents=True)
    (tmp_path / "index.html").write_text("x")
    (tmp_path / "res" / "deep" / "app.js").write_text("y")
    found = {rel for _, rel in iter_files(tmp_path)}
    assert found == {"index.html", "res/deep/app.js"}


def test_iter_files_returns_nothing_for_an_empty_directory(tmp_path):
    assert list(iter_files(tmp_path)) == []
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pytest tests/test_file_utils.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'app.utils.file_utils'`

- [ ] **Step 3: Write `app/utils/file_utils.py`**

```python
"""Filesystem helpers. Every job path derives from the integer job id."""

from __future__ import annotations

import shutil
from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path

SOURCE_FILENAME = "source.pptx"
PARTIAL_SUFFIX = ".part"


@dataclass(frozen=True)
class JobPaths:
    root: Path
    input_dir: Path
    working_dir: Path
    output_dir: Path
    log_dir: Path

    @property
    def source_pptx(self) -> Path:
        return self.input_dir / SOURCE_FILENAME

    @property
    def partial_pptx(self) -> Path:
        return self.input_dir / (SOURCE_FILENAME + PARTIAL_SUFFIX)


def build_job_paths(temp_root: str | Path, job_id: int) -> JobPaths:
    """Build the directory layout for a job.

    ``job_id`` must be an ``int``. Requiring an integer is what makes path
    traversal impossible: there is no string for a caller to smuggle
    separators through.
    """
    if isinstance(job_id, bool) or not isinstance(job_id, int):
        raise TypeError(f"job_id must be an int, got {type(job_id).__name__}")
    if job_id <= 0:
        raise ValueError(f"job_id must be positive, got {job_id}")

    root = Path(temp_root) / str(job_id)
    return JobPaths(
        root=root,
        input_dir=root / "input",
        working_dir=root / "working",
        output_dir=root / "output",
        log_dir=root / "logs",
    )


def create_job_dirs(paths: JobPaths) -> None:
    for directory in (
        paths.root,
        paths.input_dir,
        paths.working_dir,
        paths.output_dir,
        paths.log_dir,
    ):
        directory.mkdir(parents=True, exist_ok=True)


def assert_within(base: Path, candidate: Path) -> Path:
    """Return ``candidate`` resolved, or raise if it escapes ``base``."""
    base_resolved = Path(base).resolve()
    candidate_resolved = Path(candidate).resolve()
    if base_resolved != candidate_resolved and base_resolved not in candidate_resolved.parents:
        raise ValueError(f"path {candidate_resolved} is outside {base_resolved}")
    return candidate_resolved


def free_disk_gb(path: str | Path) -> float:
    """Free space in GiB on the volume containing ``path``.

    Walks up to the nearest existing ancestor, so it works before the job
    directory has been created.
    """
    target = Path(path)
    while not target.exists() and target != target.parent:
        target = target.parent
    return shutil.disk_usage(target).free / (1024**3)


def iter_files(root: Path) -> Iterator[tuple[Path, str]]:
    """Yield ``(absolute_path, posix_relative_path)`` for every file under root.

    Relative paths use forward slashes because they become URL and object-key
    components, which are POSIX-style regardless of the host OS.
    """
    root = Path(root)
    for entry in sorted(root.rglob("*")):
        if entry.is_file():
            yield entry, entry.relative_to(root).as_posix()
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `pytest tests/test_file_utils.py -v`
Expected: PASS, 11 tests.

- [ ] **Step 5: Write the failing test `tests/test_validators.py`**

```python
import pytest

from app.exceptions import DownloadError
from app.utils.validators import assert_looks_like_pptx, validate_download_url

PPTX_MAGIC = b"PK\x03\x04"


def test_https_url_is_accepted():
    url = "https://storage.example.com/input/biology.pptx"
    assert validate_download_url(url, []) == url


def test_http_url_is_accepted():
    url = "http://storage.example.com/input/biology.pptx"
    assert validate_download_url(url, []) == url


@pytest.mark.parametrize(
    "url",
    ["file:///C:/Windows/win.ini", "ftp://example.com/a.pptx", "//example.com/a.pptx", ""],
)
def test_non_http_schemes_are_rejected(url):
    with pytest.raises(DownloadError, match="scheme"):
        validate_download_url(url, [])


def test_url_without_a_host_is_rejected():
    with pytest.raises(DownloadError, match="host"):
        validate_download_url("https:///input/a.pptx", [])


def test_empty_allowlist_permits_any_host():
    assert validate_download_url("https://anything.example/a.pptx", [])


def test_allowlist_permits_a_listed_host():
    url = "https://cdn.example.com/a.pptx"
    assert validate_download_url(url, ["cdn.example.com"]) == url


def test_allowlist_rejects_an_unlisted_host():
    with pytest.raises(DownloadError, match="not allowed"):
        validate_download_url("https://evil.example/a.pptx", ["cdn.example.com"])


def test_allowlist_matching_is_case_insensitive():
    url = "https://CDN.Example.COM/a.pptx"
    assert validate_download_url(url, ["cdn.example.com"]) == url


def test_error_message_does_not_leak_the_query_string():
    signed = "https://evil.example/a.pptx?X-Amz-Signature=deadbeef"
    with pytest.raises(DownloadError) as exc:
        validate_download_url(signed, ["cdn.example.com"])
    assert "deadbeef" not in str(exc.value)


def test_valid_pptx_passes(tmp_path):
    path = tmp_path / "a.pptx"
    path.write_bytes(PPTX_MAGIC + b"rest of the zip")
    assert_looks_like_pptx(path) is None


def test_missing_file_is_rejected(tmp_path):
    with pytest.raises(DownloadError, match="does not exist"):
        assert_looks_like_pptx(tmp_path / "missing.pptx")


def test_empty_file_is_rejected(tmp_path):
    path = tmp_path / "a.pptx"
    path.write_bytes(b"")
    with pytest.raises(DownloadError, match="empty"):
        assert_looks_like_pptx(path)


def test_html_error_page_saved_as_pptx_is_rejected(tmp_path):
    path = tmp_path / "a.pptx"
    path.write_bytes(b"<!DOCTYPE html><html><body>403 Forbidden</body></html>")
    with pytest.raises(DownloadError, match="not a valid PowerPoint"):
        assert_looks_like_pptx(path)
```

- [ ] **Step 6: Run the test to verify it fails**

Run: `pytest tests/test_validators.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'app.utils.validators'`

- [ ] **Step 7: Write `app/utils/validators.py`**

```python
"""Input validation for untrusted values."""

from __future__ import annotations

from pathlib import Path
from urllib.parse import urlsplit

from app.core.logging_config import scrub_url
from app.exceptions import DownloadError

ALLOWED_SCHEMES = frozenset({"http", "https"})
OOXML_MAGIC = b"PK\x03\x04"


def validate_download_url(url: str, allowed_hosts: list[str]) -> str:
    """Validate an input PPT URL.

    Restricts the scheme to http/https and, when ``allowed_hosts`` is
    non-empty, restricts the host. An empty allowlist permits any host; the
    resulting SSRF exposure is documented in the README.
    """
    parts = urlsplit(url or "")
    if parts.scheme.lower() not in ALLOWED_SCHEMES:
        raise DownloadError(
            f"unsupported URL scheme {parts.scheme!r}; only http and https are allowed",
            user_message="The PPT URL scheme is not supported.",
        )
    if not parts.hostname:
        raise DownloadError(
            "URL has no host",
            user_message="The PPT URL has no host.",
        )
    if allowed_hosts:
        permitted = {host.lower() for host in allowed_hosts}
        if parts.hostname.lower() not in permitted:
            raise DownloadError(
                f"host {parts.hostname!r} is not allowed by DOWNLOAD_ALLOWED_HOSTS "
                f"(url: {scrub_url(url)})",
                user_message="The PPT URL host is not permitted.",
            )
    return url


def assert_looks_like_pptx(path: Path) -> None:
    """Raise unless ``path`` is a non-empty file starting with ZIP magic bytes.

    A ``.pptx`` is an OOXML ZIP archive. This check catches the common case of
    a storage provider returning an HTML error page with a 200 status.
    """
    if not path.exists():
        raise DownloadError(
            f"downloaded file does not exist: {path.name}",
            user_message="The downloaded file is missing.",
        )
    if path.stat().st_size == 0:
        raise DownloadError(
            f"downloaded file is empty: {path.name}",
            user_message="The downloaded file is empty.",
        )
    with path.open("rb") as handle:
        header = handle.read(len(OOXML_MAGIC))
    if header != OOXML_MAGIC:
        raise DownloadError(
            f"downloaded file is not a valid PowerPoint package "
            f"(header {header!r}, expected {OOXML_MAGIC!r})",
            user_message="The downloaded file is not a valid PowerPoint file.",
        )
```

- [ ] **Step 8: Run the test to verify it passes**

Run: `pytest tests/test_validators.py -v`
Expected: PASS, 16 tests.

- [ ] **Step 9: Commit**

```bash
git add app/utils/file_utils.py app/utils/validators.py tests/test_file_utils.py tests/test_validators.py
git commit -m "feat: path safety and input validation utilities"
```

---

### Task 5: Database access layer

**Files:**
- Create: `app/core/constants.py`
- Create: `app/integrations/database.py`
- Create: `tests/fakes.py`
- Test: `tests/test_database.py`

**Interfaces:**
- Consumes: `app.core.config.Settings`, `app.exceptions.DatabaseError`.
- Produces:
  - `app.core.constants.JobStatus` — `IntEnum` with `PENDING = 0`, `PROCESSING = 1`, `COMPLETED = 2`, `FAILED = 3`.
  - `app.core.constants.ClaimOutcome` — `StrEnum` with `CLAIMED`, `NOT_FOUND`, `ALREADY_PROCESSING`, `ALREADY_COMPLETED`, `ALREADY_FAILED`, `MATERIAL_MISMATCH`, `WORKER_BUSY`.
  - `app.core.constants.ERROR_MESSAGE_MAX_CHARS = 60000`
  - `app.integrations.database.JobRow` — frozen dataclass: `id: int`, `material_id: int`, `material_name: str`, `institution_name: str`, `ppt_file_url: str`, `status: JobStatus`, `error_message: str | None`.
  - `app.integrations.database.create_db_engine(settings) -> Engine`
  - `app.integrations.database.JobRepository(engine)` with `get_job(job_id) -> JobRow | None`, `claim_pending_job(job_id) -> bool`, `mark_processing(job_id) -> bool`, `mark_completed(job_id) -> bool`, `mark_failed(job_id, error_message) -> bool`, `get_status(job_id) -> JobStatus | None`, `check_connectivity() -> bool`.
  - `tests.fakes.FakeConnection`, `tests.fakes.FakeEngine`, `tests.fakes.FakeJobRepository`.

- [ ] **Step 1: Write `app/core/constants.py`**

```python
"""Shared enumerations. Values match the existing database schema exactly."""

from __future__ import annotations

from enum import IntEnum, StrEnum


class JobStatus(IntEnum):
    PENDING = 0
    PROCESSING = 1
    COMPLETED = 2
    FAILED = 3


class ClaimOutcome(StrEnum):
    """Why a submission was or was not accepted. Returned to the caller."""

    CLAIMED = "claimed"
    NOT_FOUND = "not_found"
    ALREADY_PROCESSING = "already_processing"
    ALREADY_COMPLETED = "already_completed"
    ALREADY_FAILED = "already_failed"
    MATERIAL_MISMATCH = "material_mismatch"
    WORKER_BUSY = "worker_busy"


# error_message is a MySQL TEXT column (65,535 bytes). Leave headroom for
# multi-byte characters so a long traceback cannot fail the UPDATE.
ERROR_MESSAGE_MAX_CHARS = 60000
```

- [ ] **Step 2: Write the failing test `tests/test_database.py`**

```python
import pytest

from app.core.constants import ERROR_MESSAGE_MAX_CHARS, JobStatus
from app.exceptions import DatabaseError
from app.integrations.database import JobRepository, JobRow
from tests.fakes import FakeEngine


def _row(**overrides):
    values = {
        "id": 101,
        "material_id": 5001,
        "material_name": "Introduction to Biology",
        "institution_name": "ABC College",
        "ppt_file_url": "https://storage.example.com/input/biology.pptx",
        "status": 0,
        "error_message": None,
    }
    values.update(overrides)
    return values


def _sql_of(call):
    return " ".join(str(call.statement).split())


def test_claim_sql_matches_the_specified_statement():
    engine = FakeEngine(rowcount=1)
    assert JobRepository(engine).claim_pending_job(101) is True

    sql = _sql_of(engine.calls[0])
    assert sql == (
        "UPDATE ppt_automation_jobs "
        "SET status = 1, updated_at = NOW() "
        "WHERE id = :job_id AND status = 0"
    )
    assert engine.calls[0].params == {"job_id": 101}


def test_claim_returns_false_when_no_row_was_updated():
    engine = FakeEngine(rowcount=0)
    assert JobRepository(engine).claim_pending_job(101) is False


def test_claim_uses_a_bound_parameter_not_interpolation():
    engine = FakeEngine(rowcount=1)
    JobRepository(engine).claim_pending_job(101)
    assert "101" not in _sql_of(engine.calls[0])


def test_mark_processing_delegates_to_the_same_atomic_claim():
    engine = FakeEngine(rowcount=1)
    assert JobRepository(engine).mark_processing(101) is True
    assert _sql_of(engine.calls[0]).endswith("AND status = 0")


def test_mark_completed_only_transitions_from_processing():
    engine = FakeEngine(rowcount=1)
    assert JobRepository(engine).mark_completed(101) is True

    sql = _sql_of(engine.calls[0])
    assert "SET status = 2" in sql
    assert "error_message = NULL" in sql
    assert "AND status = 1" in sql
    assert engine.calls[0].params == {"job_id": 101}


def test_mark_completed_returns_false_if_the_job_was_not_processing():
    engine = FakeEngine(rowcount=0)
    assert JobRepository(engine).mark_completed(101) is False


def test_mark_failed_records_the_message_and_allows_pending_or_processing():
    engine = FakeEngine(rowcount=1)
    assert JobRepository(engine).mark_failed(101, "download: timed out") is True

    sql = _sql_of(engine.calls[0])
    assert "SET status = 3" in sql
    assert "error_message = :error_message" in sql
    assert "AND status IN (0, 1)" in sql
    assert engine.calls[0].params == {
        "job_id": 101,
        "error_message": "download: timed out",
    }


def test_mark_failed_truncates_an_oversized_message():
    engine = FakeEngine(rowcount=1)
    JobRepository(engine).mark_failed(101, "x" * (ERROR_MESSAGE_MAX_CHARS + 500))
    assert len(engine.calls[0].params["error_message"]) == ERROR_MESSAGE_MAX_CHARS


def test_mark_failed_never_writes_completed_over_a_finished_job():
    engine = FakeEngine(rowcount=1)
    JobRepository(engine).mark_failed(101, "boom")
    sql = _sql_of(engine.calls[0])
    assert "status IN (0, 1)" in sql
    assert "status IN (0, 1, 2)" not in sql


def test_get_job_maps_a_row_to_the_dataclass():
    engine = FakeEngine(rows=[_row()])
    job = JobRepository(engine).get_job(101)
    assert job == JobRow(
        id=101,
        material_id=5001,
        material_name="Introduction to Biology",
        institution_name="ABC College",
        ppt_file_url="https://storage.example.com/input/biology.pptx",
        status=JobStatus.PENDING,
        error_message=None,
    )


def test_get_job_returns_none_when_absent():
    engine = FakeEngine(rows=[])
    assert JobRepository(engine).get_job(999) is None


def test_get_job_selects_named_columns_not_a_wildcard():
    engine = FakeEngine(rows=[_row()])
    JobRepository(engine).get_job(101)
    assert "SELECT *" not in _sql_of(engine.calls[0])


def test_get_status_returns_the_enum():
    engine = FakeEngine(rows=[{"status": 2}])
    assert JobRepository(engine).get_status(101) is JobStatus.COMPLETED


def test_get_status_returns_none_when_absent():
    engine = FakeEngine(rows=[])
    assert JobRepository(engine).get_status(999) is None


def test_check_connectivity_true_on_success():
    engine = FakeEngine(rows=[{"1": 1}])
    assert JobRepository(engine).check_connectivity() is True


def test_check_connectivity_false_on_failure():
    engine = FakeEngine(raise_on_execute=True)
    assert JobRepository(engine).check_connectivity() is False


def test_driver_errors_are_wrapped_in_databaseerror():
    engine = FakeEngine(raise_on_execute=True)
    with pytest.raises(DatabaseError):
        JobRepository(engine).claim_pending_job(101)
```

- [ ] **Step 3: Write `tests/fakes.py`**

These fakes are shared by later tasks. `FakeEngine` records the exact statement and parameters, which is how the claim SQL is verified without a MySQL server.

```python
"""Test doubles shared across the suite."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from sqlalchemy.exc import OperationalError

from app.core.constants import ClaimOutcome, JobStatus
from app.integrations.database import JobRow


@dataclass
class RecordedCall:
    statement: Any
    params: dict[str, Any] | None


class FakeResult:
    def __init__(self, rowcount: int = 0, rows: list[dict] | None = None) -> None:
        self.rowcount = rowcount
        self._rows = list(rows or [])

    def mappings(self) -> "FakeResult":
        return self

    def first(self) -> dict | None:
        return self._rows[0] if self._rows else None


class FakeConnection:
    def __init__(self, engine: "FakeEngine") -> None:
        self._engine = engine

    def __enter__(self) -> "FakeConnection":
        return self

    def __exit__(self, *exc_info: object) -> None:
        return None

    def execute(self, statement: Any, params: dict | None = None) -> FakeResult:
        self._engine.calls.append(RecordedCall(statement=statement, params=params))
        if self._engine.raise_on_execute:
            raise OperationalError("SELECT 1", {}, Exception("connection refused"))
        return FakeResult(rowcount=self._engine.rowcount, rows=self._engine.rows)


class FakeEngine:
    """Stands in for a SQLAlchemy Engine and records every statement."""

    def __init__(
        self,
        rowcount: int = 0,
        rows: list[dict] | None = None,
        raise_on_execute: bool = False,
    ) -> None:
        self.rowcount = rowcount
        self.rows = rows or []
        self.raise_on_execute = raise_on_execute
        self.calls: list[RecordedCall] = []

    def connect(self) -> FakeConnection:
        return FakeConnection(self)

    def begin(self) -> FakeConnection:
        return FakeConnection(self)


@dataclass
class FakeJobRepository:
    """In-memory repository with the same contract as JobRepository."""

    jobs: dict[int, JobRow] = field(default_factory=dict)
    connectivity: bool = True
    claim_calls: list[int] = field(default_factory=list)
    completed: list[int] = field(default_factory=list)
    failures: list[tuple[int, str]] = field(default_factory=list)

    def add(self, job: JobRow) -> None:
        self.jobs[job.id] = job

    def get_job(self, job_id: int) -> JobRow | None:
        return self.jobs.get(job_id)

    def claim_pending_job(self, job_id: int) -> bool:
        self.claim_calls.append(job_id)
        job = self.jobs.get(job_id)
        if job is None or job.status is not JobStatus.PENDING:
            return False
        self.jobs[job_id] = replace_status(job, JobStatus.PROCESSING)
        return True

    def mark_processing(self, job_id: int) -> bool:
        return self.claim_pending_job(job_id)

    def mark_completed(self, job_id: int) -> bool:
        job = self.jobs.get(job_id)
        if job is None or job.status is not JobStatus.PROCESSING:
            return False
        self.jobs[job_id] = replace_status(job, JobStatus.COMPLETED)
        self.completed.append(job_id)
        return True

    def mark_failed(self, job_id: int, error_message: str) -> bool:
        job = self.jobs.get(job_id)
        if job is None or job.status not in (JobStatus.PENDING, JobStatus.PROCESSING):
            return False
        self.jobs[job_id] = replace_status(job, JobStatus.FAILED, error_message)
        self.failures.append((job_id, error_message))
        return True

    def get_status(self, job_id: int) -> JobStatus | None:
        job = self.jobs.get(job_id)
        return job.status if job else None

    def check_connectivity(self) -> bool:
        return self.connectivity


def replace_status(
    job: JobRow, status: JobStatus, error_message: str | None = None
) -> JobRow:
    from dataclasses import replace

    return replace(job, status=status, error_message=error_message)


def make_job_row(**overrides: Any) -> JobRow:
    values: dict[str, Any] = {
        "id": 101,
        "material_id": 5001,
        "material_name": "Introduction to Biology",
        "institution_name": "ABC College",
        "ppt_file_url": "https://storage.example.com/input/biology.pptx",
        "status": JobStatus.PENDING,
        "error_message": None,
    }
    values.update(overrides)
    return JobRow(**values)


__all__ = [
    "ClaimOutcome",
    "FakeConnection",
    "FakeEngine",
    "FakeJobRepository",
    "FakeResult",
    "RecordedCall",
    "make_job_row",
]
```

- [ ] **Step 4: Run the test to verify it fails**

Run: `pytest tests/test_database.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'app.integrations.database'`

- [ ] **Step 5: Write `app/integrations/database.py`**

`CLAIM_SQL` is written as a literal to match the specified statement exactly. Do not refactor it into the generic status updater — the whole duplicate-protection guarantee rests on this statement, and it should be readable and greppable in one piece.

```python
"""MySQL access via SQLAlchemy Core. No ORM, no DDL, bound parameters only."""

from __future__ import annotations

import logging
from dataclasses import dataclass

from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine
from sqlalchemy.exc import SQLAlchemyError

from app.core.config import Settings
from app.core.constants import ERROR_MESSAGE_MAX_CHARS, JobStatus
from app.exceptions import DatabaseError

logger = logging.getLogger(__name__)

# The single-claim guarantee. The affected row count is the only authority on
# whether this worker owns the job.
CLAIM_SQL = text(
    """
    UPDATE ppt_automation_jobs
    SET status = 1,
        updated_at = NOW()
    WHERE id = :job_id
      AND status = 0
    """
)

COMPLETE_SQL = text(
    """
    UPDATE ppt_automation_jobs
    SET status = 2,
        error_message = NULL,
        updated_at = NOW()
    WHERE id = :job_id
      AND status = 1
    """
)

FAIL_SQL = text(
    """
    UPDATE ppt_automation_jobs
    SET status = 3,
        error_message = :error_message,
        updated_at = NOW()
    WHERE id = :job_id
      AND status IN (0, 1)
    """
)

SELECT_JOB_SQL = text(
    """
    SELECT id,
           material_id,
           material_name,
           institution_name,
           ppt_file_url,
           status,
           error_message
    FROM ppt_automation_jobs
    WHERE id = :job_id
    """
)

SELECT_STATUS_SQL = text(
    """
    SELECT status
    FROM ppt_automation_jobs
    WHERE id = :job_id
    """
)

PING_SQL = text("SELECT 1")


@dataclass(frozen=True)
class JobRow:
    id: int
    material_id: int
    material_name: str
    institution_name: str
    ppt_file_url: str
    status: JobStatus
    error_message: str | None


def create_db_engine(settings: Settings) -> Engine:
    """Build the engine. The driver lives in DATABASE_URL and is swappable."""
    if not settings.database_url:
        raise DatabaseError(
            "DATABASE_URL is not configured",
            user_message="The service is not configured to reach the database.",
        )
    return create_engine(
        settings.database_url,
        pool_pre_ping=True,
        pool_recycle=1800,
        pool_size=5,
        max_overflow=2,
        future=True,
    )


class JobRepository:
    """Every read and write against ppt_automation_jobs."""

    def __init__(self, engine: Engine) -> None:
        self._engine = engine

    def claim_pending_job(self, job_id: int) -> bool:
        """Atomically move PENDING -> PROCESSING. True means we own the job."""
        try:
            with self._engine.begin() as conn:
                result = conn.execute(CLAIM_SQL, {"job_id": job_id})
                return result.rowcount == 1
        except SQLAlchemyError as exc:
            logger.error(
                "claim failed", extra={"job_id": job_id, "stage": "claim"}, exc_info=True
            )
            raise DatabaseError(f"could not claim job {job_id}: {exc}") from exc

    def mark_processing(self, job_id: int) -> bool:
        """Named transition into PROCESSING.

        Delegates to :meth:`claim_pending_job` because PENDING -> PROCESSING is
        the only legal way to enter PROCESSING; having two statements that could
        drift apart would be a duplicate-protection bug waiting to happen.
        """
        return self.claim_pending_job(job_id)

    def mark_completed(self, job_id: int) -> bool:
        """PROCESSING -> COMPLETED, clearing any previous error message."""
        try:
            with self._engine.begin() as conn:
                result = conn.execute(COMPLETE_SQL, {"job_id": job_id})
                return result.rowcount == 1
        except SQLAlchemyError as exc:
            logger.error(
                "mark_completed failed",
                extra={"job_id": job_id, "stage": "complete"},
                exc_info=True,
            )
            raise DatabaseError(f"could not complete job {job_id}: {exc}") from exc

    def mark_failed(self, job_id: int, error_message: str) -> bool:
        """PENDING or PROCESSING -> FAILED. Never overwrites COMPLETED."""
        try:
            with self._engine.begin() as conn:
                result = conn.execute(
                    FAIL_SQL,
                    {
                        "job_id": job_id,
                        "error_message": error_message[:ERROR_MESSAGE_MAX_CHARS],
                    },
                )
                return result.rowcount == 1
        except SQLAlchemyError as exc:
            logger.error(
                "mark_failed failed",
                extra={"job_id": job_id, "stage": "database"},
                exc_info=True,
            )
            raise DatabaseError(f"could not fail job {job_id}: {exc}") from exc

    def get_job(self, job_id: int) -> JobRow | None:
        try:
            with self._engine.connect() as conn:
                row = conn.execute(SELECT_JOB_SQL, {"job_id": job_id}).mappings().first()
        except SQLAlchemyError as exc:
            logger.error(
                "get_job failed",
                extra={"job_id": job_id, "stage": "database"},
                exc_info=True,
            )
            raise DatabaseError(f"could not read job {job_id}: {exc}") from exc
        if row is None:
            return None
        return JobRow(
            id=row["id"],
            material_id=row["material_id"],
            material_name=row["material_name"],
            institution_name=row["institution_name"],
            ppt_file_url=row["ppt_file_url"],
            status=JobStatus(row["status"]),
            error_message=row["error_message"],
        )

    def get_status(self, job_id: int) -> JobStatus | None:
        try:
            with self._engine.connect() as conn:
                row = (
                    conn.execute(SELECT_STATUS_SQL, {"job_id": job_id})
                    .mappings()
                    .first()
                )
        except SQLAlchemyError as exc:
            logger.error(
                "get_status failed",
                extra={"job_id": job_id, "stage": "database"},
                exc_info=True,
            )
            raise DatabaseError(f"could not read status for job {job_id}: {exc}") from exc
        return JobStatus(row["status"]) if row is not None else None

    def check_connectivity(self) -> bool:
        """Used by /health. Never raises."""
        try:
            with self._engine.connect() as conn:
                conn.execute(PING_SQL)
            return True
        except SQLAlchemyError:
            logger.warning("database connectivity check failed", exc_info=True)
            return False
```

- [ ] **Step 6: Run the test to verify it passes**

Run: `pytest tests/test_database.py -v`
Expected: PASS, 17 tests.

- [ ] **Step 7: Commit**

```bash
git add app/core/constants.py app/integrations/database.py tests/fakes.py tests/test_database.py
git commit -m "feat: job repository with atomic claim and guarded status transitions"
```

---

### Task 6: Single-slot concurrency guard

**Files:**
- Create: `app/core/job_slot.py`
- Test: `tests/test_job_slot.py`

**Interfaces:**
- Consumes: nothing.
- Produces: `app.core.job_slot.JobSlot` with `try_acquire() -> bool`, `release() -> None`, `is_busy` property, `current_job_id: int | None`, and `try_acquire(job_id: int)` recording the holder.

A `threading.Lock` is used deliberately instead of an `asyncio.Lock`. `asyncio.Lock` has no non-blocking acquire, so the only way to express "reject if busy" is to check `locked()` and then acquire — two operations, and another coroutine can be scheduled between them. `threading.Lock.acquire(blocking=False)` is a single atomic operation, and it is also callable from the COM worker thread.

- [ ] **Step 1: Write the failing test `tests/test_job_slot.py`**

```python
import threading

import pytest

from app.core.job_slot import JobSlot


def test_first_acquire_succeeds():
    slot = JobSlot()
    assert slot.try_acquire(101) is True
    assert slot.is_busy is True
    assert slot.current_job_id == 101


def test_second_acquire_is_refused_while_held():
    slot = JobSlot()
    slot.try_acquire(101)
    assert slot.try_acquire(102) is False
    assert slot.current_job_id == 101


def test_release_frees_the_slot():
    slot = JobSlot()
    slot.try_acquire(101)
    slot.release()
    assert slot.is_busy is False
    assert slot.current_job_id is None
    assert slot.try_acquire(102) is True


def test_release_when_not_held_does_not_raise():
    JobSlot().release()


def test_only_one_of_many_threads_acquires():
    slot = JobSlot()
    winners: list[int] = []
    barrier = threading.Barrier(20)

    def contend(job_id: int) -> None:
        barrier.wait()
        if slot.try_acquire(job_id):
            winners.append(job_id)

    threads = [threading.Thread(target=contend, args=(i,)) for i in range(20)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()

    assert len(winners) == 1


def test_context_manager_releases_on_success():
    slot = JobSlot()
    with slot.hold(101) as acquired:
        assert acquired is True
    assert slot.is_busy is False


def test_context_manager_releases_on_exception():
    slot = JobSlot()
    with pytest.raises(RuntimeError):
        with slot.hold(101):
            raise RuntimeError("boom")
    assert slot.is_busy is False


def test_context_manager_yields_false_when_busy_and_does_not_release_the_holder():
    slot = JobSlot()
    slot.try_acquire(101)
    with slot.hold(102) as acquired:
        assert acquired is False
    assert slot.is_busy is True
    assert slot.current_job_id == 101
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pytest tests/test_job_slot.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'app.core.job_slot'`

- [ ] **Step 3: Write `app/core/job_slot.py`**

```python
"""The process-wide guarantee that only one PPT is processed at a time."""

from __future__ import annotations

import logging
import threading
from collections.abc import Iterator
from contextlib import contextmanager

logger = logging.getLogger(__name__)


class JobSlot:
    """A single occupancy slot with a non-blocking acquire.

    Uses ``threading.Lock`` rather than ``asyncio.Lock``: the latter has no
    non-blocking acquire, so expressing "reject if busy" would require
    checking ``locked()`` and then acquiring, and another coroutine can run
    between those two steps. ``threading.Lock.acquire(blocking=False)`` is
    atomic and is also usable from the COM worker thread.
    """

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._current_job_id: int | None = None

    def try_acquire(self, job_id: int) -> bool:
        acquired = self._lock.acquire(blocking=False)
        if acquired:
            self._current_job_id = job_id
        else:
            logger.info(
                "worker busy, refusing job",
                extra={"job_id": job_id, "holder_job_id": self._current_job_id},
            )
        return acquired

    def release(self) -> None:
        self._current_job_id = None
        try:
            self._lock.release()
        except RuntimeError:
            # Releasing an unheld lock is not an error worth propagating; it
            # means a caller's finally block ran without a matching acquire.
            logger.debug("release called while the slot was not held")

    @property
    def is_busy(self) -> bool:
        return self._lock.locked()

    @property
    def current_job_id(self) -> int | None:
        return self._current_job_id

    @contextmanager
    def hold(self, job_id: int) -> Iterator[bool]:
        """Yield whether the slot was acquired; release only if it was."""
        acquired = self.try_acquire(job_id)
        try:
            yield acquired
        finally:
            if acquired:
                self.release()
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `pytest tests/test_job_slot.py -v`
Expected: PASS, 8 tests.

- [ ] **Step 5: Commit**

```bash
git add app/core/job_slot.py tests/test_job_slot.py
git commit -m "feat: single-slot concurrency guard with atomic non-blocking acquire"
```

---

### Task 7: API key authentication

**Files:**
- Create: `app/core/security.py`
- Test: `tests/test_security.py`

**Interfaces:**
- Consumes: `app.core.config.Settings`, `app.core.config.get_settings`.
- Produces: `app.core.security.API_KEY_HEADER = "X-Worker-API-Key"`; `app.core.security.verify_api_key(provided: str | None, expected: str) -> bool`; `app.core.security.require_api_key` — a FastAPI dependency raising `HTTPException(401)`.

`secrets.compare_digest` is used rather than `==` so comparison time does not depend on how many leading characters matched.

- [ ] **Step 1: Write the failing test `tests/test_security.py`**

```python
import pytest
from fastapi import HTTPException

from app.core.config import Settings
from app.core.security import API_KEY_HEADER, require_api_key, verify_api_key


def _settings(key: str = "correct-key") -> Settings:
    return Settings(worker_api_key=key)


def test_header_name():
    assert API_KEY_HEADER == "X-Worker-API-Key"


def test_correct_key_verifies():
    assert verify_api_key("correct-key", "correct-key") is True


def test_wrong_key_is_rejected():
    assert verify_api_key("wrong-key", "correct-key") is False


def test_missing_key_is_rejected():
    assert verify_api_key(None, "correct-key") is False


def test_empty_provided_key_is_rejected():
    assert verify_api_key("", "correct-key") is False


def test_empty_expected_key_rejects_everything():
    assert verify_api_key("anything", "") is False
    assert verify_api_key("", "") is False


def test_non_ascii_key_does_not_raise():
    assert verify_api_key("ké", "correct-key") is False


def test_dependency_passes_with_a_valid_key():
    assert require_api_key(x_worker_api_key="correct-key", settings=_settings()) is None


def test_dependency_raises_401_without_a_key():
    with pytest.raises(HTTPException) as exc:
        require_api_key(x_worker_api_key=None, settings=_settings())
    assert exc.value.status_code == 401


def test_dependency_raises_401_with_a_wrong_key():
    with pytest.raises(HTTPException) as exc:
        require_api_key(x_worker_api_key="nope", settings=_settings())
    assert exc.value.status_code == 401


def test_401_detail_does_not_echo_the_expected_or_provided_key():
    with pytest.raises(HTTPException) as exc:
        require_api_key(x_worker_api_key="nope", settings=_settings())
    detail = str(exc.value.detail)
    assert "correct-key" not in detail
    assert "nope" not in detail
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pytest tests/test_security.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'app.core.security'`

- [ ] **Step 3: Write `app/core/security.py`**

```python
"""Worker API key authentication."""

from __future__ import annotations

import logging
import secrets
from typing import Annotated

from fastapi import Depends, Header, HTTPException, status

from app.core.config import Settings, get_settings

logger = logging.getLogger(__name__)

API_KEY_HEADER = "X-Worker-API-Key"


def verify_api_key(provided: str | None, expected: str) -> bool:
    """Constant-time comparison. An unset expected key rejects everything."""
    if not provided or not expected:
        return False
    return secrets.compare_digest(provided.encode("utf-8"), expected.encode("utf-8"))


def require_api_key(
    x_worker_api_key: Annotated[str | None, Header(alias=API_KEY_HEADER)] = None,
    settings: Annotated[Settings, Depends(get_settings)] = None,  # type: ignore[assignment]
) -> None:
    """FastAPI dependency. Raises 401 before any other work happens."""
    if not verify_api_key(x_worker_api_key, settings.worker_api_key):
        logger.warning(
            "rejected request with an invalid worker API key",
            extra={"key_present": bool(x_worker_api_key)},
        )
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid or missing worker API key.",
        )
    return None
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `pytest tests/test_security.py -v`
Expected: PASS, 11 tests.

- [ ] **Step 5: Commit**

```bash
git add app/core/security.py tests/test_security.py
git commit -m "feat: constant-time worker API key authentication"
```

---

### Task 8: Request and response schemas

**Files:**
- Create: `app/schemas/jobs.py`
- Create: `app/schemas/health.py`
- Test: `tests/test_schemas.py`

**Interfaces:**
- Consumes: `app.core.constants.JobStatus`, `app.core.constants.ClaimOutcome`.
- Produces:
  - `app.schemas.jobs.ProcessJobRequest` — `job_id: int`, `material_id: int`, `material_name: str`, `institution_name: str`, `ppt_file_url: str`.
  - `app.schemas.jobs.ProcessJobResponse` — `job_id: int`, `accepted: bool`, `outcome: ClaimOutcome`, `message: str`.
  - `app.schemas.jobs.JobStatusResponse` — `job_id: int`, `material_id: int`, `status: int`, `status_name: str`, `error_message: str | None`.
  - `app.schemas.health.CheckResult` — `name: str`, `ok: bool`, `detail: str`.
  - `app.schemas.health.HealthResponse` — `status: Literal["ok", "degraded"]`, `checks: list[CheckResult]`.

- [ ] **Step 1: Write the failing test `tests/test_schemas.py`**

```python
import pytest
from pydantic import ValidationError

from app.core.constants import ClaimOutcome, JobStatus
from app.schemas.health import CheckResult, HealthResponse
from app.schemas.jobs import JobStatusResponse, ProcessJobRequest, ProcessJobResponse

VALID = {
    "job_id": 101,
    "material_id": 5001,
    "material_name": "Introduction to Biology",
    "institution_name": "ABC College",
    "ppt_file_url": "https://storage.example.com/input/biology.pptx",
}


def test_valid_request_parses():
    request = ProcessJobRequest(**VALID)
    assert request.job_id == 101
    assert request.material_id == 5001


@pytest.mark.parametrize("missing", list(VALID))
def test_every_field_is_required(missing):
    payload = {k: v for k, v in VALID.items() if k != missing}
    with pytest.raises(ValidationError):
        ProcessJobRequest(**payload)


def test_job_id_must_be_positive():
    with pytest.raises(ValidationError):
        ProcessJobRequest(**{**VALID, "job_id": 0})


def test_material_id_must_be_positive():
    with pytest.raises(ValidationError):
        ProcessJobRequest(**{**VALID, "material_id": -1})


def test_job_id_rejects_a_non_numeric_string():
    with pytest.raises(ValidationError):
        ProcessJobRequest(**{**VALID, "job_id": "abc"})


def test_names_are_length_limited_to_the_column_width():
    with pytest.raises(ValidationError):
        ProcessJobRequest(**{**VALID, "material_name": "x" * 501})
    with pytest.raises(ValidationError):
        ProcessJobRequest(**{**VALID, "institution_name": "x" * 501})


def test_blank_names_are_rejected():
    with pytest.raises(ValidationError):
        ProcessJobRequest(**{**VALID, "material_name": "   "})


def test_names_are_whitespace_stripped():
    request = ProcessJobRequest(**{**VALID, "material_name": "  Biology  "})
    assert request.material_name == "Biology"


def test_response_carries_the_outcome():
    response = ProcessJobResponse(
        job_id=101, accepted=False, outcome=ClaimOutcome.WORKER_BUSY, message="busy"
    )
    assert response.outcome == "worker_busy"


def test_status_response_includes_the_readable_name():
    response = JobStatusResponse(
        job_id=101,
        material_id=5001,
        status=JobStatus.COMPLETED,
        status_name="COMPLETED",
        error_message=None,
    )
    assert response.status == 2
    assert response.status_name == "COMPLETED"


def test_health_response_shape():
    response = HealthResponse(
        status="degraded",
        checks=[CheckResult(name="database", ok=False, detail="unreachable")],
    )
    assert response.status == "degraded"
    assert response.checks[0].name == "database"


def test_health_status_rejects_an_unknown_value():
    with pytest.raises(ValidationError):
        HealthResponse(status="on fire", checks=[])
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pytest tests/test_schemas.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'app.schemas.jobs'`

- [ ] **Step 3: Write `app/schemas/jobs.py`**

The `max_length=500` values mirror the `VARCHAR(500)` columns, so an oversized name is rejected at the edge with a 422 rather than truncated or failing mid-pipeline.

```python
"""Request and response models for the job endpoints."""

from __future__ import annotations

from pydantic import BaseModel, Field, field_validator

from app.core.constants import ClaimOutcome


class ProcessJobRequest(BaseModel):
    job_id: int = Field(gt=0, description="Primary key in ppt_automation_jobs")
    material_id: int = Field(gt=0)
    material_name: str = Field(min_length=1, max_length=500)
    institution_name: str = Field(min_length=1, max_length=500)
    ppt_file_url: str = Field(min_length=1)

    @field_validator("material_name", "institution_name")
    @classmethod
    def _strip_and_require_content(cls, value: str) -> str:
        stripped = value.strip()
        if not stripped:
            raise ValueError("must not be blank")
        return stripped


class ProcessJobResponse(BaseModel):
    job_id: int
    accepted: bool
    outcome: ClaimOutcome
    message: str


class JobStatusResponse(BaseModel):
    job_id: int
    material_id: int
    status: int
    status_name: str
    error_message: str | None = None
```

- [ ] **Step 4: Write `app/schemas/health.py`**

```python
"""Response models for the health endpoint."""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel


class CheckResult(BaseModel):
    name: str
    ok: bool
    detail: str = ""


class HealthResponse(BaseModel):
    status: Literal["ok", "degraded"]
    checks: list[CheckResult]
```

- [ ] **Step 5: Run the test to verify it passes**

Run: `pytest tests/test_schemas.py -v`
Expected: PASS, 17 tests.

- [ ] **Step 6: Commit**

```bash
git add app/schemas/jobs.py app/schemas/health.py tests/test_schemas.py
git commit -m "feat: request and response schemas with edge validation"
```

---

### Task 9: Job service — claim orchestration and outcome mapping

**Files:**
- Create: `app/services/job_service.py`
- Test: `tests/test_job_service.py`

**Interfaces:**
- Consumes: `app.core.constants.ClaimOutcome`, `app.core.constants.JobStatus`, `app.core.job_slot.JobSlot`, `app.integrations.database.JobRow`, `app.schemas.jobs.ProcessJobRequest`, `app.schemas.jobs.JobStatusResponse`, `app.exceptions.DatabaseError`.
- Produces:
  - `app.services.job_service.ClaimResult` — frozen dataclass: `outcome: ClaimOutcome`, `accepted: bool`, `http_status: int`, `message: str`, `job: JobRow | None`.
  - `app.services.job_service.JobPipeline` — `Protocol` with `async def run(self, job: JobRow) -> None`.
  - `app.services.job_service.JobService(repository, slot)` with `claim(request: ProcessJobRequest) -> ClaimResult`, `release_slot() -> None`, `get_status(job_id: int) -> JobStatusResponse | None`.

The ordering inside `claim` is the whole point of this task. The slot is acquired **before** the database is touched, so a busy worker returns `423` without ever writing `status = 1`. Claiming first and rejecting afterwards would leave an orphaned `PROCESSING` row on every collision, needing a manual reset.

- [ ] **Step 1: Write the failing test `tests/test_job_service.py`**

```python
import pytest

from app.core.constants import ClaimOutcome, JobStatus
from app.core.job_slot import JobSlot
from app.exceptions import DatabaseError
from app.schemas.jobs import ProcessJobRequest
from app.services.job_service import JobService
from tests.fakes import FakeJobRepository, make_job_row

REQUEST = ProcessJobRequest(
    job_id=101,
    material_id=5001,
    material_name="Introduction to Biology",
    institution_name="ABC College",
    ppt_file_url="https://storage.example.com/input/biology.pptx",
)


def _service(*jobs):
    repo = FakeJobRepository()
    for job in jobs:
        repo.add(job)
    return JobService(repository=repo, slot=JobSlot()), repo


def test_pending_job_is_claimed():
    service, repo = _service(make_job_row())
    result = service.claim(REQUEST)
    assert result.accepted is True
    assert result.outcome is ClaimOutcome.CLAIMED
    assert result.http_status == 202
    assert result.job is not None
    assert repo.get_status(101) is JobStatus.PROCESSING


def test_claimed_job_leaves_the_slot_held_for_the_background_task():
    service, _ = _service(make_job_row())
    service.claim(REQUEST)
    assert service.slot.is_busy is True


def test_busy_worker_returns_423_and_never_touches_the_database():
    service, repo = _service(make_job_row())
    service.slot.try_acquire(999)

    result = service.claim(REQUEST)

    assert result.accepted is False
    assert result.outcome is ClaimOutcome.WORKER_BUSY
    assert result.http_status == 423
    assert repo.claim_calls == []
    assert repo.get_status(101) is JobStatus.PENDING


def test_missing_job_returns_not_found():
    service, _ = _service()
    result = service.claim(REQUEST)
    assert result.outcome is ClaimOutcome.NOT_FOUND
    assert result.http_status == 409
    assert result.accepted is False


def test_already_processing_returns_409():
    service, _ = _service(make_job_row(status=JobStatus.PROCESSING))
    result = service.claim(REQUEST)
    assert result.outcome is ClaimOutcome.ALREADY_PROCESSING
    assert result.http_status == 409


def test_already_completed_returns_409():
    service, _ = _service(make_job_row(status=JobStatus.COMPLETED))
    assert service.claim(REQUEST).outcome is ClaimOutcome.ALREADY_COMPLETED


def test_already_failed_returns_409():
    service, _ = _service(make_job_row(status=JobStatus.FAILED))
    assert service.claim(REQUEST).outcome is ClaimOutcome.ALREADY_FAILED


@pytest.mark.parametrize(
    "status",
    [JobStatus.PROCESSING, JobStatus.COMPLETED, JobStatus.FAILED],
)
def test_slot_is_released_when_the_claim_is_refused(status):
    service, _ = _service(make_job_row(status=status))
    service.claim(REQUEST)
    assert service.slot.is_busy is False


def test_slot_is_released_when_the_job_is_missing():
    service, _ = _service()
    service.claim(REQUEST)
    assert service.slot.is_busy is False


def test_material_id_mismatch_fails_the_job_and_is_refused():
    service, repo = _service(make_job_row(material_id=9999))
    result = service.claim(REQUEST)

    assert result.outcome is ClaimOutcome.MATERIAL_MISMATCH
    assert result.http_status == 409
    assert result.accepted is False
    assert repo.get_status(101) is JobStatus.FAILED
    assert service.slot.is_busy is False


def test_material_id_mismatch_records_a_stage_prefixed_error():
    service, repo = _service(make_job_row(material_id=9999))
    service.claim(REQUEST)
    _, message = repo.failures[0]
    assert message.startswith("claim:")
    assert "9999" in message and "5001" in message


def test_database_error_during_claim_releases_the_slot_and_propagates():
    class ExplodingRepository(FakeJobRepository):
        def claim_pending_job(self, job_id: int) -> bool:
            raise DatabaseError("connection refused")

    service = JobService(repository=ExplodingRepository(), slot=JobSlot())
    with pytest.raises(DatabaseError):
        service.claim(REQUEST)
    assert service.slot.is_busy is False


def test_release_slot_is_safe_to_call_twice():
    service, _ = _service(make_job_row())
    service.claim(REQUEST)
    service.release_slot()
    service.release_slot()
    assert service.slot.is_busy is False


def test_get_status_returns_a_response_for_a_known_job():
    service, _ = _service(make_job_row(status=JobStatus.FAILED, error_message="boom"))
    response = service.get_status(101)
    assert response is not None
    assert response.status == 3
    assert response.status_name == "FAILED"
    assert response.error_message == "boom"


def test_get_status_returns_none_for_an_unknown_job():
    service, _ = _service()
    assert service.get_status(999) is None
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pytest tests/test_job_service.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'app.services.job_service'`

- [ ] **Step 3: Write `app/services/job_service.py`**

```python
"""Claim orchestration and status reporting.

This module owns the two duplicate-protection guarantees: the in-process slot
and the conditional database update. It performs no PowerPoint or filesystem
work and imports nothing platform-specific.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Protocol

from app.core.constants import ClaimOutcome, JobStatus
from app.core.job_slot import JobSlot
from app.integrations.database import JobRow
from app.schemas.jobs import JobStatusResponse, ProcessJobRequest

logger = logging.getLogger(__name__)

_STATUS_TO_OUTCOME = {
    JobStatus.PROCESSING: ClaimOutcome.ALREADY_PROCESSING,
    JobStatus.COMPLETED: ClaimOutcome.ALREADY_COMPLETED,
    JobStatus.FAILED: ClaimOutcome.ALREADY_FAILED,
}


@dataclass(frozen=True)
class ClaimResult:
    outcome: ClaimOutcome
    accepted: bool
    http_status: int
    message: str
    job: JobRow | None = None


class JobPipeline(Protocol):
    """The processing pipeline. Implemented in Plan 2 by JobRunner."""

    async def run(self, job: JobRow) -> None: ...


class JobService:
    def __init__(self, repository, slot: JobSlot) -> None:
        self._repository = repository
        self.slot = slot

    def claim(self, request: ProcessJobRequest) -> ClaimResult:
        """Attempt to take ownership of a job.

        On success the slot is left held; the caller's background task must
        call :meth:`release_slot` in a finally block. On every refusal the
        slot is released here.
        """
        if not self.slot.try_acquire(request.job_id):
            # Deliberately before any database access: a busy worker must not
            # write status=1, so the job stays PENDING for the next cron tick.
            return ClaimResult(
                outcome=ClaimOutcome.WORKER_BUSY,
                accepted=False,
                http_status=423,
                message=(
                    "Another PPT is being processed. The job was left pending; "
                    "retry on the next cycle."
                ),
            )

        try:
            claimed = self._repository.claim_pending_job(request.job_id)
            if not claimed:
                return self._refuse(self._classify_unclaimable(request.job_id))

            job = self._repository.get_job(request.job_id)
            if job is None:
                logger.error(
                    "job disappeared immediately after a successful claim",
                    extra={"job_id": request.job_id, "stage": "claim"},
                )
                return self._refuse(
                    ClaimResult(
                        outcome=ClaimOutcome.NOT_FOUND,
                        accepted=False,
                        http_status=409,
                        message="Job could not be read after claiming.",
                    )
                )

            if job.material_id != request.material_id:
                detail = (
                    f"claim: material_id mismatch, database has {job.material_id} "
                    f"but the request said {request.material_id}"
                )
                logger.error(
                    "material_id mismatch, refusing to process",
                    extra={
                        "job_id": request.job_id,
                        "material_id": request.material_id,
                        "db_material_id": job.material_id,
                        "stage": "claim",
                    },
                )
                self._repository.mark_failed(request.job_id, detail)
                return self._refuse(
                    ClaimResult(
                        outcome=ClaimOutcome.MATERIAL_MISMATCH,
                        accepted=False,
                        http_status=409,
                        message="The request's material_id does not match the job.",
                    )
                )

            logger.info(
                "job claimed",
                extra={
                    "job_id": job.id,
                    "material_id": job.material_id,
                    "stage": "claim",
                },
            )
            return ClaimResult(
                outcome=ClaimOutcome.CLAIMED,
                accepted=True,
                http_status=202,
                message="Job claimed and queued for processing.",
                job=job,
            )
        except Exception:
            self.release_slot()
            logger.error(
                "claim raised, slot released",
                extra={"job_id": request.job_id, "stage": "claim"},
                exc_info=True,
            )
            raise

    def _refuse(self, result: ClaimResult) -> ClaimResult:
        self.release_slot()
        logger.info(
            "job submission refused",
            extra={"job_id": result.job.id if result.job else None, "outcome": result.outcome},
        )
        return result

    def _classify_unclaimable(self, job_id: int) -> ClaimResult:
        status = self._repository.get_status(job_id)
        if status is None:
            return ClaimResult(
                outcome=ClaimOutcome.NOT_FOUND,
                accepted=False,
                http_status=409,
                message="No job exists with that id.",
            )
        outcome = _STATUS_TO_OUTCOME.get(status)
        if outcome is None:
            # status was PENDING but the UPDATE matched nothing: another
            # process claimed it between our UPDATE and this SELECT.
            outcome = ClaimOutcome.ALREADY_PROCESSING
        return ClaimResult(
            outcome=outcome,
            accepted=False,
            http_status=409,
            message=f"Job is not claimable ({outcome}).",
        )

    def release_slot(self) -> None:
        self.slot.release()

    def get_status(self, job_id: int) -> JobStatusResponse | None:
        job = self._repository.get_job(job_id)
        if job is None:
            return None
        return JobStatusResponse(
            job_id=job.id,
            material_id=job.material_id,
            status=int(job.status),
            status_name=job.status.name,
            error_message=job.error_message,
        )
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `pytest tests/test_job_service.py -v`
Expected: PASS, 17 tests.

- [ ] **Step 5: Commit**

```bash
git add app/services/job_service.py tests/test_job_service.py
git commit -m "feat: claim orchestration with slot-before-database ordering"
```

---

### Task 10: Health service and endpoint

**Files:**
- Create: `app/services/health_service.py`
- Create: `app/api/health.py`
- Test: `tests/test_health_service.py`

**Interfaces:**
- Consumes: `app.core.config.Settings`, `app.schemas.health.CheckResult`, `app.schemas.health.HealthResponse`, `app.utils.file_utils.free_disk_gb`.
- Produces:
  - `app.services.health_service.HealthService(settings, repository, powerpoint_probe=None, ispring_probe=None)` with `check() -> HealthResponse`.
  - A probe is any `Callable[[], CheckResult]`.
  - `app.api.health.router` — `APIRouter` exposing `GET /health`.

The endpoint returns HTTP 200 even when degraded, with a per-check breakdown. A monitor that only sees a 503 knows something is wrong but not what; the breakdown is what makes the page actionable. In this plan the PowerPoint and iSpring probes report `ok=False, detail="not wired up yet"` because those services arrive in Plan 2 — which is honest, since an unconfigured worker genuinely cannot publish.

- [ ] **Step 1: Write the failing test `tests/test_health_service.py`**

```python
from app.core.config import Settings
from app.schemas.health import CheckResult
from app.services.health_service import HealthService
from tests.fakes import FakeJobRepository


def _settings(tmp_path, **overrides):
    values = {
        "worker_api_key": "k",
        "temp_root": str(tmp_path / "jobs"),
        "log_root": str(tmp_path / "logs"),
        "min_free_disk_gb": 1,
    }
    values.update(overrides)
    return Settings(**values)


def _ok(name):
    return lambda: CheckResult(name=name, ok=True, detail="available")


def _service(tmp_path, repository=None, **overrides):
    return HealthService(
        settings=_settings(tmp_path, **overrides),
        repository=repository or FakeJobRepository(),
        powerpoint_probe=_ok("powerpoint"),
        ispring_probe=_ok("ispring"),
    )


def _by_name(response):
    return {check.name: check for check in response.checks}


def test_all_checks_passing_reports_ok(tmp_path):
    response = _service(tmp_path).check()
    assert response.status == "ok"


def test_every_expected_check_is_present(tmp_path):
    checks = _by_name(_service(tmp_path).check())
    assert set(checks) == {
        "fastapi",
        "powerpoint",
        "ispring",
        "temp_dir_writable",
        "disk_space",
        "database",
    }


def test_fastapi_check_is_always_ok(tmp_path):
    assert _by_name(_service(tmp_path).check())["fastapi"].ok is True


def test_temp_dir_check_creates_and_probes_the_directory(tmp_path):
    service = _service(tmp_path)
    check = _by_name(service.check())["temp_dir_writable"]
    assert check.ok is True
    assert (tmp_path / "jobs").is_dir()


def test_temp_dir_probe_leaves_no_file_behind(tmp_path):
    _service(tmp_path).check()
    assert list((tmp_path / "jobs").iterdir()) == []


def test_unwritable_temp_dir_degrades(tmp_path):
    blocker = tmp_path / "blocker"
    blocker.write_text("not a directory")
    response = _service(tmp_path, temp_root=str(blocker / "jobs")).check()
    assert response.status == "degraded"
    assert _by_name(response)["temp_dir_writable"].ok is False


def test_low_disk_space_degrades(tmp_path):
    response = _service(tmp_path, min_free_disk_gb=10**9).check()
    assert response.status == "degraded"
    assert _by_name(response)["disk_space"].ok is False


def test_database_unreachable_degrades(tmp_path):
    repository = FakeJobRepository(connectivity=False)
    response = _service(tmp_path, repository=repository).check()
    assert response.status == "degraded"
    assert _by_name(response)["database"].ok is False


def test_failing_powerpoint_probe_degrades(tmp_path):
    service = HealthService(
        settings=_settings(tmp_path),
        repository=FakeJobRepository(),
        powerpoint_probe=lambda: CheckResult(
            name="powerpoint", ok=False, detail="not installed"
        ),
        ispring_probe=_ok("ispring"),
    )
    assert service.check().status == "degraded"


def test_a_raising_probe_degrades_rather_than_crashing(tmp_path):
    def boom() -> CheckResult:
        raise RuntimeError("COM blew up")

    service = HealthService(
        settings=_settings(tmp_path),
        repository=FakeJobRepository(),
        powerpoint_probe=boom,
        ispring_probe=_ok("ispring"),
    )
    response = service.check()
    assert response.status == "degraded"
    assert _by_name(response)["powerpoint"].ok is False


def test_default_probes_report_not_wired_up(tmp_path):
    service = HealthService(settings=_settings(tmp_path), repository=FakeJobRepository())
    checks = _by_name(service.check())
    assert checks["powerpoint"].ok is False
    assert checks["ispring"].ok is False


def test_no_check_detail_leaks_a_secret(tmp_path):
    settings = _settings(tmp_path, worker_api_key="super-secret", database_url="mysql+pymysql://u:pw@h/db")
    service = HealthService(settings=settings, repository=FakeJobRepository())
    blob = " ".join(check.detail for check in service.check().checks)
    assert "super-secret" not in blob
    assert "pw" not in blob
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `pytest tests/test_health_service.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'app.services.health_service'`

- [ ] **Step 3: Write `app/services/health_service.py`**

```python
"""Health checks. Reports detail without revealing configuration secrets."""

from __future__ import annotations

import logging
import os
from collections.abc import Callable
from pathlib import Path

from app.core.config import Settings
from app.schemas.health import CheckResult, HealthResponse
from app.utils.file_utils import free_disk_gb

logger = logging.getLogger(__name__)

Probe = Callable[[], CheckResult]


def _not_wired_up(name: str) -> Probe:
    def probe() -> CheckResult:
        return CheckResult(
            name=name,
            ok=False,
            detail="not wired up yet; provided by the processing pipeline",
        )

    return probe


class HealthService:
    def __init__(
        self,
        settings: Settings,
        repository,
        powerpoint_probe: Probe | None = None,
        ispring_probe: Probe | None = None,
    ) -> None:
        self._settings = settings
        self._repository = repository
        self._powerpoint_probe = powerpoint_probe or _not_wired_up("powerpoint")
        self._ispring_probe = ispring_probe or _not_wired_up("ispring")

    def check(self) -> HealthResponse:
        checks = [
            CheckResult(name="fastapi", ok=True, detail="running"),
            self._safe(self._powerpoint_probe, "powerpoint"),
            self._safe(self._ispring_probe, "ispring"),
            self._safe(self._check_temp_dir, "temp_dir_writable"),
            self._safe(self._check_disk_space, "disk_space"),
            self._safe(self._check_database, "database"),
        ]
        status = "ok" if all(check.ok for check in checks) else "degraded"
        return HealthResponse(status=status, checks=checks)

    def _safe(self, probe: Probe, name: str) -> CheckResult:
        """Run a probe, converting any exception into a failed check.

        A probe that raises must not take the health endpoint down with it —
        that is precisely when the endpoint is most needed.
        """
        try:
            return probe()
        except Exception as exc:
            logger.warning("health probe failed", extra={"check": name}, exc_info=True)
            return CheckResult(name=name, ok=False, detail=f"probe raised: {type(exc).__name__}")

    def _check_temp_dir(self) -> CheckResult:
        temp_root = Path(self._settings.temp_root)
        probe_file = temp_root / f".health-probe-{os.getpid()}"
        try:
            temp_root.mkdir(parents=True, exist_ok=True)
            probe_file.write_text("ok", encoding="utf-8")
        except OSError as exc:
            return CheckResult(
                name="temp_dir_writable",
                ok=False,
                detail=f"not writable: {type(exc).__name__}",
            )
        finally:
            try:
                probe_file.unlink(missing_ok=True)
            except OSError:
                logger.warning("could not remove the health probe file", exc_info=True)
        return CheckResult(name="temp_dir_writable", ok=True, detail="writable")

    def _check_disk_space(self) -> CheckResult:
        free_gb = free_disk_gb(self._settings.temp_root)
        required = self._settings.min_free_disk_gb
        return CheckResult(
            name="disk_space",
            ok=free_gb >= required,
            detail=f"{free_gb:.1f} GiB free, {required} GiB required",
        )

    def _check_database(self) -> CheckResult:
        reachable = self._repository.check_connectivity()
        return CheckResult(
            name="database",
            ok=reachable,
            detail="reachable" if reachable else "unreachable",
        )
```

- [ ] **Step 4: Write `app/api/health.py`**

```python
"""GET /health."""

from __future__ import annotations

from fastapi import APIRouter, Depends

from app.api.deps import get_health_service
from app.schemas.health import HealthResponse
from app.services.health_service import HealthService

router = APIRouter(tags=["health"])


@router.get("/health", response_model=HealthResponse)
def health(service: HealthService = Depends(get_health_service)) -> HealthResponse:
    """Always returns 200. Read ``status`` and ``checks`` for the verdict."""
    return service.check()
```

- [ ] **Step 5: Run the test to verify it passes**

Run: `pytest tests/test_health_service.py -v`
Expected: PASS, 12 tests. (`app/api/health.py` is not exercised until Task 11 supplies `app/api/deps.py`.)

- [ ] **Step 6: Commit**

```bash
git add app/services/health_service.py app/api/health.py tests/test_health_service.py
git commit -m "feat: health service with per-check breakdown and safe probes"
```

---

### Task 11: Application factory, dependencies, and job endpoints

**Files:**
- Create: `app/api/deps.py`
- Create: `app/api/jobs.py`
- Create: `app/main.py`
- Create: `run.py`
- Test: `tests/conftest.py`
- Test: `tests/test_api_jobs.py`
- Test: `tests/test_api_health.py`

**Interfaces:**
- Consumes: everything from Tasks 1–10.
- Produces:
  - `app.main.create_app(settings=None, repository=None, pipeline=None, slot=None, powerpoint_probe=None, ispring_probe=None) -> FastAPI`
  - `app.main.app` — the module-level instance uvicorn loads.
  - `app.api.deps.get_job_service(request) -> JobService`, `get_pipeline(request) -> JobPipeline`, `get_health_service(request) -> HealthService`, `get_app_settings(request) -> Settings`.
  - `app.api.jobs.router` — `POST /jobs/process`, `GET /jobs/{job_id}`.

`create_app` takes its collaborators as optional arguments so tests inject fakes by constructing an app rather than patching globals. When they are omitted the real engine, repository, and slot are built from settings.

- [ ] **Step 1: Write `app/api/deps.py`**

```python
"""Request-scoped accessors for objects held on application state."""

from __future__ import annotations

from fastapi import Request

from app.core.config import Settings
from app.services.health_service import HealthService
from app.services.job_service import JobPipeline, JobService


def get_app_settings(request: Request) -> Settings:
    return request.app.state.settings


def get_job_service(request: Request) -> JobService:
    return request.app.state.job_service


def get_pipeline(request: Request) -> JobPipeline:
    return request.app.state.pipeline


def get_health_service(request: Request) -> HealthService:
    return request.app.state.health_service
```

- [ ] **Step 2: Write the failing test `tests/conftest.py`**

```python
import asyncio

import pytest
from httpx import ASGITransport, AsyncClient

from app.core.config import Settings
from app.core.job_slot import JobSlot
from app.integrations.database import JobRow
from app.main import create_app
from tests.fakes import FakeJobRepository, make_job_row

API_KEY = "test-worker-key"

VALID_PAYLOAD = {
    "job_id": 101,
    "material_id": 5001,
    "material_name": "Introduction to Biology",
    "institution_name": "ABC College",
    "ppt_file_url": "https://storage.example.com/input/biology.pptx",
}


class RecordingPipeline:
    """Records the jobs it was asked to run."""

    def __init__(self) -> None:
        self.runs: list[JobRow] = []
        self.started = asyncio.Event()

    async def run(self, job: JobRow) -> None:
        self.runs.append(job)
        self.started.set()


class BlockingPipeline(RecordingPipeline):
    """Holds the slot until ``release`` is set, simulating a long publish."""

    def __init__(self) -> None:
        super().__init__()
        self.release = asyncio.Event()

    async def run(self, job: JobRow) -> None:
        self.runs.append(job)
        self.started.set()
        await self.release.wait()


@pytest.fixture
def settings(tmp_path) -> Settings:
    return Settings(
        app_env="development",
        worker_api_key=API_KEY,
        temp_root=str(tmp_path / "jobs"),
        log_root=str(tmp_path / "logs"),
        min_free_disk_gb=0,
    )


@pytest.fixture
def repository() -> FakeJobRepository:
    repo = FakeJobRepository()
    repo.add(make_job_row())
    return repo


@pytest.fixture
def pipeline() -> RecordingPipeline:
    return RecordingPipeline()


@pytest.fixture
def app(settings, repository, pipeline):
    return create_app(
        settings=settings,
        repository=repository,
        pipeline=pipeline,
        slot=JobSlot(),
    )


@pytest.fixture
async def client(app):
    transport = ASGITransport(app=app)
    async with AsyncClient(transport=transport, base_url="http://worker") as async_client:
        yield async_client


@pytest.fixture
def auth_headers() -> dict[str, str]:
    return {"X-Worker-API-Key": API_KEY}
```

- [ ] **Step 3: Write the failing test `tests/test_api_jobs.py`**

```python
import asyncio

import pytest
from httpx import ASGITransport, AsyncClient

from app.core.constants import JobStatus
from app.core.job_slot import JobSlot
from app.main import create_app
from tests.conftest import VALID_PAYLOAD, BlockingPipeline
from tests.fakes import FakeJobRepository, make_job_row


async def _wait_for(condition, timeout: float = 2.0) -> None:
    async with asyncio.timeout(timeout):
        while not condition():
            await asyncio.sleep(0.01)


async def test_valid_submission_is_accepted(client, auth_headers, repository, pipeline):
    response = await client.post("/jobs/process", json=VALID_PAYLOAD, headers=auth_headers)

    assert response.status_code == 202
    body = response.json()
    assert body["accepted"] is True
    assert body["outcome"] == "claimed"
    assert body["job_id"] == 101
    assert repository.get_status(101) is JobStatus.PROCESSING

    await _wait_for(lambda: pipeline.runs)
    assert pipeline.runs[0].id == 101


async def test_missing_api_key_is_401(client):
    response = await client.post("/jobs/process", json=VALID_PAYLOAD)
    assert response.status_code == 401


async def test_wrong_api_key_is_401(client, repository):
    response = await client.post(
        "/jobs/process", json=VALID_PAYLOAD, headers={"X-Worker-API-Key": "nope"}
    )
    assert response.status_code == 401
    assert repository.get_status(101) is JobStatus.PENDING


async def test_auth_is_checked_before_the_body_is_validated(client, repository):
    response = await client.post("/jobs/process", json={"job_id": "garbage"})
    assert response.status_code == 401
    assert repository.claim_calls == []


@pytest.mark.parametrize("missing", list(VALID_PAYLOAD))
async def test_missing_field_is_422(client, auth_headers, missing):
    payload = {k: v for k, v in VALID_PAYLOAD.items() if k != missing}
    response = await client.post("/jobs/process", json=payload, headers=auth_headers)
    assert response.status_code == 422


async def test_unknown_job_is_409_not_found(client, auth_headers):
    payload = {**VALID_PAYLOAD, "job_id": 999, "material_id": 5001}
    response = await client.post("/jobs/process", json=payload, headers=auth_headers)
    assert response.status_code == 409
    assert response.json()["outcome"] == "not_found"


@pytest.mark.parametrize(
    ("status", "expected"),
    [
        (JobStatus.PROCESSING, "already_processing"),
        (JobStatus.COMPLETED, "already_completed"),
        (JobStatus.FAILED, "already_failed"),
    ],
)
async def test_unclaimable_statuses_are_409(settings, pipeline, status, expected):
    repository = FakeJobRepository()
    repository.add(make_job_row(status=status))
    app = create_app(
        settings=settings, repository=repository, pipeline=pipeline, slot=JobSlot()
    )
    async with AsyncClient(
        transport=ASGITransport(app=app), base_url="http://worker"
    ) as client:
        response = await client.post(
            "/jobs/process",
            json=VALID_PAYLOAD,
            headers={"X-Worker-API-Key": settings.worker_api_key},
        )
    assert response.status_code == 409
    assert response.json()["outcome"] == expected
    assert pipeline.runs == []


async def test_resubmitting_the_same_job_is_refused_the_second_time(
    client, auth_headers, repository
):
    first = await client.post("/jobs/process", json=VALID_PAYLOAD, headers=auth_headers)
    second = await client.post("/jobs/process", json=VALID_PAYLOAD, headers=auth_headers)

    assert first.status_code == 202
    assert second.status_code in (409, 423)
    assert repository.claim_calls.count(101) == 1


async def test_concurrent_submissions_yield_exactly_one_acceptance(settings):
    """The duplicate-protection test. Two requests race for one slot."""
    repository = FakeJobRepository()
    repository.add(make_job_row())
    repository.add(make_job_row(id=102, material_id=5002))
    pipeline = BlockingPipeline()
    app = create_app(
        settings=settings, repository=repository, pipeline=pipeline, slot=JobSlot()
    )
    headers = {"X-Worker-API-Key": settings.worker_api_key}

    async with AsyncClient(
        transport=ASGITransport(app=app), base_url="http://worker"
    ) as client:
        first, second = await asyncio.gather(
            client.post("/jobs/process", json=VALID_PAYLOAD, headers=headers),
            client.post(
                "/jobs/process",
                json={**VALID_PAYLOAD, "job_id": 102, "material_id": 5002},
                headers=headers,
            ),
        )
        codes = sorted([first.status_code, second.status_code])
        assert codes == [202, 423]

        busy = first if first.status_code == 423 else second
        assert busy.json()["outcome"] == "worker_busy"

        await _wait_for(lambda: pipeline.runs)
        assert len(pipeline.runs) == 1
        pipeline.release.set()

    # The refused job was never claimed, so the next cron tick can retry it.
    refused_id = 102 if pipeline.runs[0].id == 101 else 101
    assert repository.get_status(refused_id) is JobStatus.PENDING


async def test_busy_worker_does_not_touch_the_database(settings):
    repository = FakeJobRepository()
    repository.add(make_job_row())
    pipeline = BlockingPipeline()
    slot = JobSlot()
    slot.try_acquire(999)
    app = create_app(
        settings=settings, repository=repository, pipeline=pipeline, slot=slot
    )

    async with AsyncClient(
        transport=ASGITransport(app=app), base_url="http://worker"
    ) as client:
        response = await client.post(
            "/jobs/process",
            json=VALID_PAYLOAD,
            headers={"X-Worker-API-Key": settings.worker_api_key},
        )

    assert response.status_code == 423
    assert repository.claim_calls == []
    assert repository.get_status(101) is JobStatus.PENDING


async def test_slot_is_released_after_the_pipeline_finishes(client, auth_headers, app, pipeline):
    await client.post("/jobs/process", json=VALID_PAYLOAD, headers=auth_headers)
    await _wait_for(lambda: pipeline.runs)
    await _wait_for(lambda: not app.state.slot.is_busy)
    assert app.state.slot.is_busy is False


async def test_slot_is_released_when_the_pipeline_raises(settings, repository):
    class ExplodingPipeline:
        async def run(self, job):
            raise RuntimeError("publish exploded")

    app = create_app(
        settings=settings,
        repository=repository,
        pipeline=ExplodingPipeline(),
        slot=JobSlot(),
    )
    async with AsyncClient(
        transport=ASGITransport(app=app), base_url="http://worker"
    ) as client:
        response = await client.post(
            "/jobs/process",
            json=VALID_PAYLOAD,
            headers={"X-Worker-API-Key": settings.worker_api_key},
        )
        assert response.status_code == 202
        await _wait_for(lambda: not app.state.slot.is_busy)

    assert app.state.slot.is_busy is False


async def test_get_status_returns_the_job(client, auth_headers):
    response = await client.get("/jobs/101", headers=auth_headers)
    assert response.status_code == 200
    body = response.json()
    assert body["status"] == 0
    assert body["status_name"] == "PENDING"


async def test_get_status_unknown_job_is_404(client, auth_headers):
    response = await client.get("/jobs/999", headers=auth_headers)
    assert response.status_code == 404


async def test_get_status_requires_the_api_key(client):
    response = await client.get("/jobs/101")
    assert response.status_code == 401
```

- [ ] **Step 4: Write the failing test `tests/test_api_health.py`**

```python
async def test_health_returns_200(client):
    response = await client.get("/health")
    assert response.status_code == 200


async def test_health_needs_no_api_key(client):
    assert (await client.get("/health")).status_code == 200


async def test_health_reports_a_status_and_checks(client):
    body = (await client.get("/health")).json()
    assert body["status"] in ("ok", "degraded")
    names = {check["name"] for check in body["checks"]}
    assert {"fastapi", "temp_dir_writable", "disk_space", "database"} <= names


async def test_health_never_returns_the_api_key(client):
    assert "test-worker-key" not in (await client.get("/health")).text
```

- [ ] **Step 5: Run both tests to verify they fail**

Run: `pytest tests/test_api_jobs.py tests/test_api_health.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'app.main'`

- [ ] **Step 6: Write `app/api/jobs.py`**

The background task is created with `asyncio.create_task` and its reference stored on `app.state.tasks`, because a task with no live reference can be garbage collected before it completes. The `finally` block releases the slot, so it is freed whether the pipeline succeeds, fails, or is cancelled.

```python
"""POST /jobs/process and GET /jobs/{job_id}."""

from __future__ import annotations

import asyncio
import logging

from fastapi import APIRouter, Depends, HTTPException, Request, Response, status

from app.api.deps import get_job_service, get_pipeline
from app.core.logging_config import bind_job_context, clear_job_context
from app.core.security import require_api_key
from app.integrations.database import JobRow
from app.schemas.jobs import JobStatusResponse, ProcessJobRequest, ProcessJobResponse
from app.services.job_service import JobPipeline, JobService

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/jobs", tags=["jobs"], dependencies=[Depends(require_api_key)])


async def _run_and_release(
    pipeline: JobPipeline, service: JobService, job: JobRow
) -> None:
    bind_job_context(job_id=job.id, material_id=job.material_id)
    try:
        await pipeline.run(job)
    except Exception:
        logger.error(
            "pipeline raised out of the background task",
            extra={"job_id": job.id, "stage": "unexpected"},
            exc_info=True,
        )
    finally:
        service.release_slot()
        clear_job_context()


@router.post(
    "/process",
    response_model=ProcessJobResponse,
    status_code=status.HTTP_202_ACCEPTED,
)
async def process_job(
    payload: ProcessJobRequest,
    request: Request,
    response: Response,
    service: JobService = Depends(get_job_service),
    pipeline: JobPipeline = Depends(get_pipeline),
) -> ProcessJobResponse:
    """Claim the job, schedule it, and return immediately.

    The claim happens here rather than in the background task, so a 202
    response is itself proof the job was claimed exactly once.
    """
    result = service.claim(payload)
    response.status_code = result.http_status

    if not result.accepted or result.job is None:
        return ProcessJobResponse(
            job_id=payload.job_id,
            accepted=False,
            outcome=result.outcome,
            message=result.message,
        )

    task = asyncio.create_task(_run_and_release(pipeline, service, result.job))
    tasks: set[asyncio.Task] = request.app.state.tasks
    tasks.add(task)
    task.add_done_callback(tasks.discard)

    return ProcessJobResponse(
        job_id=result.job.id,
        accepted=True,
        outcome=result.outcome,
        message=result.message,
    )


@router.get("/{job_id}", response_model=JobStatusResponse)
def get_job_status(
    job_id: int, service: JobService = Depends(get_job_service)
) -> JobStatusResponse:
    status_response = service.get_status(job_id)
    if status_response is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND, detail="No job exists with that id."
        )
    return status_response
```

- [ ] **Step 7: Write `app/main.py`**

```python
"""Application factory and lifespan."""

from __future__ import annotations

import logging
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse

from app.api import health as health_api
from app.api import jobs as jobs_api
from app.core.config import Settings, get_settings
from app.core.job_slot import JobSlot
from app.core.logging_config import configure_logging
from app.exceptions import PptAutomationError
from app.integrations.database import JobRepository, create_db_engine
from app.services.health_service import HealthService
from app.services.job_service import JobPipeline, JobService

logger = logging.getLogger(__name__)


class UnconfiguredPipeline:
    """Placeholder pipeline used until Plan 2 supplies JobRunner."""

    async def run(self, job) -> None:
        raise PptAutomationError(
            "the processing pipeline is not installed on this build",
            stage="unexpected",
        )


@asynccontextmanager
async def _lifespan(app: FastAPI):
    settings: Settings = app.state.settings
    logger.info(
        "worker starting",
        extra={
            "app_env": settings.app_env,
            "storage_backend": settings.storage_backend,
            "ispring_adapter": settings.ispring_adapter,
        },
    )
    yield
    in_flight = {task for task in app.state.tasks if not task.done()}
    if in_flight:
        logger.warning(
            "shutting down with jobs in flight; they will be interrupted",
            extra={"in_flight": len(in_flight)},
        )
        for task in in_flight:
            task.cancel()
    logger.info("worker stopped")


def create_app(
    settings: Settings | None = None,
    repository=None,
    pipeline: JobPipeline | None = None,
    slot: JobSlot | None = None,
    powerpoint_probe=None,
    ispring_probe=None,
) -> FastAPI:
    """Build the app. Collaborators are injectable so tests can supply fakes."""
    settings = settings or get_settings()
    configure_logging(settings)

    if repository is None:
        repository = JobRepository(create_db_engine(settings))

    slot = slot or JobSlot()
    pipeline = pipeline or UnconfiguredPipeline()

    app = FastAPI(
        title="PPT Automation Worker",
        version="1.0.0",
        lifespan=_lifespan,
    )
    app.state.settings = settings
    app.state.repository = repository
    app.state.slot = slot
    app.state.pipeline = pipeline
    app.state.tasks = set()
    app.state.job_service = JobService(repository=repository, slot=slot)
    app.state.health_service = HealthService(
        settings=settings,
        repository=repository,
        powerpoint_probe=powerpoint_probe,
        ispring_probe=ispring_probe,
    )

    app.include_router(health_api.router)
    app.include_router(jobs_api.router)

    @app.exception_handler(PptAutomationError)
    async def _handle_domain_error(request: Request, exc: PptAutomationError):
        logger.error(
            "request failed", extra={"stage": exc.stage}, exc_info=True
        )
        return JSONResponse(
            status_code=500,
            content={"detail": exc.user_message, "stage": exc.stage},
        )

    return app


app = create_app()
```

- [ ] **Step 8: Write `run.py`**

```python
"""Development and service entry point.

``workers`` is hard-coded to 1: the single-PowerPoint guarantee is an
in-process lock, and a second worker process would create a second lock.
"""

from __future__ import annotations

import uvicorn

from app.core.config import get_settings


def main() -> None:
    settings = get_settings()
    uvicorn.run(
        "app.main:app",
        host=settings.app_host,
        port=settings.app_port,
        workers=1,
        log_config=None,
        access_log=False,
    )


if __name__ == "__main__":
    main()
```

- [ ] **Step 9: Run the full suite**

Run: `pytest -v`
Expected: PASS. All tests across every file, with no PowerPoint, iSpring, or MySQL present.

- [ ] **Step 10: Verify the app boots and answers**

```bash
WORKER_API_KEY=dev-key TEMP_ROOT=./temp LOG_ROOT=./logs python run.py &
sleep 3
curl -s localhost:8000/health
curl -s -o /dev/null -w '%{http_code}\n' -X POST localhost:8000/jobs/process \
  -H 'Content-Type: application/json' -d '{"job_id":1,"material_id":1,"material_name":"a","institution_name":"b","ppt_file_url":"https://x/y.pptx"}'
kill %1
```

Expected: `/health` returns JSON with `"status":"degraded"` (PowerPoint, iSpring, and the database are not available on a dev machine). The POST returns `401`, because no API key header was sent.

- [ ] **Step 11: Commit**

```bash
git add app/api/deps.py app/api/jobs.py app/main.py run.py tests/conftest.py tests/test_api_jobs.py tests/test_api_health.py
git commit -m "feat: app factory, job endpoints, and background task scheduling"
```

---

## Plan Self-Review

**Spec coverage.** Every intake-side requirement in the spec maps to a task: process model and layering (Task 11), COM isolation (deferred to Plan 2 but enforced by the Global Constraints), execution model and the claim sequence (Tasks 6, 9, 11), the state machine and repository methods (Task 5), path safety and download validation helpers (Task 4), the exception hierarchy (Task 2), logging and redaction (Task 3), configuration with production guards (Task 1), the health endpoint (Task 10), and the API contract (Tasks 8, 11).

**Deferred to Plan 2, deliberately:** the download implementation itself, `powerpoint_service`, `ispring_service` and its adapters, `storage_service` and the three backends, `backend_service`, `slack_service`, `cleanup_service`, `job_runner`, `windows_com`, `process_utils`, `http_client`, `scripts/probe_ispring.py`, `README.md`, and `docs/MANUAL_WINDOWS_TESTS.md`. Task 4 builds the validation helpers the downloader will use, so no work is duplicated.

**Type consistency.** `JobRow` is constructed only in `JobRepository.get_job` and `tests.fakes.make_job_row`, with identical field names. `ClaimOutcome` values are asserted as the same strings in `test_job_service.py` and `test_api_jobs.py`. `JobPipeline.run` is `async def run(self, job: JobRow) -> None` in the protocol, in both test pipelines, and in `UnconfiguredPipeline`. `HealthService`'s probe type is `Callable[[], CheckResult]` everywhere it appears.

**No placeholders.** Every step contains runnable code or an exact command. The one intentional stub, `UnconfiguredPipeline`, raises rather than silently doing nothing, and Task 11 Step 10 documents the exact `degraded` health output it produces.

