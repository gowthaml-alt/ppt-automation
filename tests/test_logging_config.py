import json
import logging

from config.settings import Settings
from utils.logging_config import (
    JobContextFilter,
    JsonFormatter,
    bind_job_context,
    clear_job_context,
    configure_logging,
    scrub_url,
    set_stage,
)


def _settings(**overrides):
    values = {
        "backend_base_url": "https://backend.example.com",
        "log_root": "/tmp/ppt-automation/logs",
    }
    values.update(overrides)
    return Settings(**values)


def _record(msg: str = "hello", **extra) -> logging.LogRecord:
    record = logging.LogRecord(
        name="worker.test",
        level=logging.INFO,
        pathname=__file__,
        lineno=1,
        msg=msg,
        args=(),
        exc_info=None,
    )
    for key, value in extra.items():
        setattr(record, key, value)
    return record


def test_scrub_url_removes_the_query_string():
    signed = "https://s3.example.com/in/a.pptx?X-Amz-Signature=deadbeef&X-Amz-Expires=900"
    assert scrub_url(signed) == "https://s3.example.com/in/a.pptx"


def test_scrub_url_removes_userinfo_credentials():
    assert scrub_url("https://user:secret@example.com/a") == "https://example.com/a"


def test_job_context_uses_job_id():
    clear_job_context()
    bind_job_context(job_id=101, material_id=5001)
    set_stage("download")
    try:
        record = _record("downloading")
        JobContextFilter().filter(record)
        payload = json.loads(JsonFormatter().format(record))
        assert payload["job_id"] == 101
        assert payload["material_id"] == 5001
        assert payload["stage"] == "download"
    finally:
        clear_job_context()


def test_configure_logging_writes_a_log_file(tmp_path):
    root = logging.getLogger()
    saved_handlers = root.handlers[:]
    saved_flag = getattr(root, "_ppt_automation_configured", None)
    root.handlers = []
    if hasattr(root, "_ppt_automation_configured"):
        del root._ppt_automation_configured
    try:
        configure_logging(_settings(log_root=str(tmp_path)))
        logging.getLogger("worker.test").info("written to disk")
        for handler in root.handlers:
            handler.flush()
        assert (tmp_path / "worker.log").exists()
    finally:
        for handler in root.handlers:
            handler.close()
        root.handlers = saved_handlers
        if saved_flag is not None:
            root._ppt_automation_configured = saved_flag
