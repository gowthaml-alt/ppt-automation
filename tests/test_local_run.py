from pathlib import Path

from config.settings import Settings
from tests.pptx_bytes import minimal_pptx_bytes
from worker.local_run import run_local_job


PPTX = minimal_pptx_bytes()


def _settings(tmp_path: Path, **overrides) -> Settings:
    values = {
        "app_env": "development",
        "backend_base_url": "http://127.0.0.1:9",
        "temp_root": str(tmp_path / "jobs"),
        "log_root": str(tmp_path / "logs"),
        "min_free_disk_gb": 0,
        "storage_backend": "local_fs",
        "local_storage_root": str(tmp_path / "cdn"),
        "local_storage_public_base_url": "http://127.0.0.1/cdn",
        "ispring_adapter": "not_configured",
        "browser_test_enabled": True,
    }
    values.update(overrides)
    return Settings(**values)


def test_local_job_copies_uploads_and_http_checks(tmp_path):
    pptx = tmp_path / "sample.pptx"
    pptx.write_bytes(PPTX)
    result = run_local_job(
        pptx,
        settings=_settings(tmp_path),
        queue_id=101,
        material_id=5001,
        skip_powerpoint=True,
        use_fake_ispring=True,
        skip_callback=True,
    )
    assert result.ok is True
    assert result.error is None
    local_index = tmp_path / "cdn" / "ppt" / "5001" / "index.html"
    assert local_index.exists()
    assert "html" in local_index.read_text(encoding="utf-8").lower()
    assert result.checked_url.startswith("http://127.0.0.1:")
    assert result.browser_ok is True


def test_local_job_fails_for_bad_pptx(tmp_path):
    pptx = tmp_path / "bad.pptx"
    pptx.write_bytes(b"<html>nope</html>")
    result = run_local_job(
        pptx,
        settings=_settings(tmp_path),
        skip_powerpoint=True,
        use_fake_ispring=True,
        skip_callback=True,
    )
    assert result.ok is False
    assert result.error is not None
    assert not (tmp_path / "cdn" / "ppt" / "1" / "index.html").exists()
