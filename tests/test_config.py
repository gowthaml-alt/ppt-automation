import pytest

from config.settings import Settings


def _base(**overrides):
    values = {
        "app_env": "development",
        "backend_base_url": "https://backend.example.com",
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
    assert s.poll_interval_seconds == 30
    assert s.max_ppt_size_mb == 500
    assert s.download_allowed_hosts == []
    assert s.backend_api_key == ""


def test_max_ppt_size_bytes_is_derived():
    s = Settings(**_base(max_ppt_size_mb=2))
    assert s.max_ppt_size_bytes == 2 * 1024 * 1024


def test_download_allowed_hosts_parses_comma_separated():
    s = Settings(**_base(download_allowed_hosts="cdn.example.com, storage.example.com"))
    assert s.download_allowed_hosts == ["cdn.example.com", "storage.example.com"]


def test_production_requires_backend_url():
    with pytest.raises(Exception, match="BACKEND_BASE_URL"):
        Settings(**_base(app_env="production", backend_base_url="", storage_backend="s3"))


def test_production_rejects_local_fs_storage():
    with pytest.raises(Exception, match="STORAGE_BACKEND"):
        Settings(**_base(app_env="production", storage_backend="local_fs"))


def test_production_allows_valid_configuration():
    s = Settings(**_base(app_env="production", storage_backend="s3"))
    assert s.is_production is True


def test_no_worker_api_key_or_database_url_fields():
    s = Settings(**_base())
    assert "database_url" not in s.model_fields
    assert "worker_api_key" not in s.model_fields
