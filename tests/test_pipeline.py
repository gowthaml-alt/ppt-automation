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
    assert "working" in deps["powerpoint"].opened[0]
    assert deps["powerpoint"].last_timeout_s == 120
    assert deps["publisher"].published == deps["powerpoint"].opened
    assert not (tmp_path / "jobs" / "101").exists()


def test_publish_uses_the_file_powerpoint_actually_has_open(tmp_path):
    powerpoint = FakePowerPoint(repaired_name="source.repaired.pptx")
    pipeline, deps = _pipeline(tmp_path, powerpoint=powerpoint)
    pipeline.process(make_job())
    assert deps["backend"].results[0]["status"] == "completed"
    assert deps["publisher"].published == [
        str(Path(powerpoint.opened[0]).with_name("source.repaired.pptx"))
    ]


def test_pipeline_opens_and_publishes_working_copy_without_changing_original(tmp_path):
    from tests.pptx_bytes import minimal_pptx_bytes
    from utils.paths import build_job_paths, create_job_dirs

    original_bytes = minimal_pptx_bytes()

    def downloader(url, paths, settings, **kwargs):
        create_job_dirs(paths)
        paths.source_pptx.write_bytes(original_bytes)
        return paths.source_pptx

    class RecordingPowerPoint:
        def __init__(self) -> None:
            self.opened: list[str] = []
            self.quit_calls = 0
            self.source_after_open = b""

        def start(self) -> None:
            return None

        def open(self, path, timeout_s=None) -> None:
            path = Path(path)
            source = path.parent.parent / "input" / "source.pptx"
            self.opened.append(str(path))
            path.write_bytes(b"repaired-on-working-copy")
            self.source_after_open = source.read_bytes()

        def close(self) -> None:
            return None

        def quit(self) -> None:
            self.quit_calls += 1

    powerpoint = RecordingPowerPoint()
    pipeline, deps = _pipeline(tmp_path, powerpoint=powerpoint, downloader=downloader)
    pipeline.process(make_job())
    paths = build_job_paths(tmp_path / "jobs", 101)
    assert powerpoint.opened == [str(paths.working_pptx)]
    assert powerpoint.source_after_open == original_bytes
    assert deps["publisher"].published == [str(paths.working_pptx)]
    assert deps["powerpoint"].quit_calls >= 1 or powerpoint.quit_calls >= 1


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
