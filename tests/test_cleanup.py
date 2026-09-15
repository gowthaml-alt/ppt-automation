from pathlib import Path

from config.settings import Settings
from utils.cleanup import CleanupService
from utils.paths import build_job_paths, create_job_dirs


def _settings(tmp_path: Path, **overrides) -> Settings:
    values = {
        "app_env": "development",
        "backend_base_url": "https://backend.example.com",
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
