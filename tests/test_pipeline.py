from pathlib import Path

import pytest

from config.settings import Settings
from tests.fakes import (
    FakeBackend,
    FakeBrowser,
    FakePowerPoint,
    FakePublisher,
    FakeSlack,
    FakeStorage,
    fake_downloader,
    make_job,
)
from utils.cleanup import CleanupService
from utils.exceptions import ISpringPublishingError, OutputUploadError
from worker.pipeline import Pipeline


def _settings(tmp_path: Path, **overrides) -> Settings:
    values = {
        "app_env": "development",
        "backend_base_url": "https://backend.example.com",
        "temp_root": str(tmp_path / "jobs"),
        "log_root": str(tmp_path / "logs"),
        "min_free_disk_gb": 0,
        "local_storage_public_base_url": "https://cdn.example.com",
    }
    values.update(overrides)
    return Settings(**values)


def _pipeline(tmp_path, **overrides):
    settings = _settings(tmp_path)
    deps = {
        "backend": FakeBackend(),
        "powerpoint": FakePowerPoint(),
        "publisher": FakePublisher(),
        "storage": FakeStorage(),
        "browser": FakeBrowser(),
        "cleanup": CleanupService(settings),
        "slack": FakeSlack(),
        "downloader": fake_downloader,
    }
    deps.update(overrides)
    return Pipeline(settings, **deps), deps


def test_success_sends_completed_callback_and_cleans_up(tmp_path):
    pipeline, deps = _pipeline(tmp_path)
    pipeline.process(make_job())
    assert deps["backend"].results[0]["status"] == "completed"
    assert deps["backend"].results[0]["iframe_url"].endswith("/ppt/5001/index.html")
    assert deps["browser"].checked
    assert deps["powerpoint"].quit_calls >= 1
    assert not (tmp_path / "jobs" / "101").exists()


def test_publish_failure_sends_failed_callback(tmp_path):
    pipeline, deps = _pipeline(tmp_path, publisher=FakePublisher(fail=True))
    with pytest.raises(ISpringPublishingError):
        pipeline.process(make_job())
    assert deps["backend"].results[0]["status"] == "failed"
    assert "iSpring" in deps["backend"].results[0]["error_message"] or deps["slack"].calls
    assert deps["storage"].uploaded == []
    assert deps["slack"].calls[0]["stage"] == "ispring_publish"


def test_upload_failure_does_not_mark_success(tmp_path):
    pipeline, deps = _pipeline(tmp_path, storage=FakeStorage(fail=True))
    with pytest.raises(OutputUploadError):
        pipeline.process(make_job())
    assert deps["backend"].results[0]["status"] == "failed"
    assert deps["browser"].checked == []


def test_callback_failure_after_upload_writes_orphan_log(tmp_path):
    pipeline, deps = _pipeline(tmp_path, backend=FakeBackend(fail_callback=True))
    with pytest.raises(Exception):
        pipeline.process(make_job())
    orphan = tmp_path / "logs" / "orphaned_outputs.jsonl"
    assert orphan.exists()
    assert "5001" in orphan.read_text(encoding="utf-8")
