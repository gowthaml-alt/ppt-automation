"""Structured JSON logging with per-job context and URL redaction."""

from __future__ import annotations

import json
import logging
import logging.handlers
from contextvars import ContextVar
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlsplit, urlunsplit

from config.settings import Settings

_queue_id: ContextVar[int | None] = ContextVar("queue_id", default=None)
_material_id: ContextVar[int | None] = ContextVar("material_id", default=None)
_stage: ContextVar[str | None] = ContextVar("stage", default=None)

_RESERVED = frozenset(
    logging.LogRecord("", 0, "", 0, "", (), None).__dict__.keys()
) | {"message", "asctime", "taskName"}


def bind_job_context(*, queue_id: int, material_id: int) -> None:
    _queue_id.set(queue_id)
    _material_id.set(material_id)


def set_stage(stage: str) -> None:
    _stage.set(stage)


def clear_job_context() -> None:
    _queue_id.set(None)
    _material_id.set(None)
    _stage.set(None)


def scrub_url(url: str) -> str:
    """Return ``url`` without its query string, fragment, or userinfo."""
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
    def filter(self, record: logging.LogRecord) -> bool:
        for attr, var in (
            ("queue_id", _queue_id),
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
    log_dir.mkdir(parents=True, exist_ok=True)
    handler = logging.FileHandler(log_dir / "job.log", encoding="utf-8")
    handler.setFormatter(JsonFormatter())
    handler.addFilter(JobContextFilter())
    return handler


def configure_logging(settings: Settings) -> None:
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
