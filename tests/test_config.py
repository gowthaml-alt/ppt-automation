import pytest

from config.settings import Settings


def _base(**overrides):
    values = {
        "app_env": "development",
        "backend_base_url": "https://api.example.com/nuSource/api/v1",
        "temp_root": "/tmp/ppt-automation/jobs",
        "log_root": "/tmp/ppt-automation/logs",
    }
    values.update(overrides)
    return values


def test_defaults_are_safe(monkeypatch):
    for key in (
        "ISPRING_ADAPTER",
        "STORAGE_BACKEND",
        "KEEP_FAILED_JOB_FILES",
        "POLL_INTERVAL_SECONDS",
        "MAX_PPT_SIZE_MB",
        "DOWNLOAD_ALLOWED_HOSTS",
        "PPT_WORKER_TOKEN",
        "BACKEND_API_KEY",
    ):
        monkeypatch.delenv(key, raising=False)
    s = Settings(_env_file=None, **_base())
    assert s.ispring_adapter == "not_configured"
    assert s.storage_backend == "local_fs"
    assert s.keep_failed_job_files is False
    assert s.poll_interval_seconds == 30
    assert s.max_ppt_size_mb == 500
    assert s.download_allowed_hosts == []
    assert "ppt_worker_token" not in s.model_fields


def test_max_ppt_size_bytes_is_derived():
    s = Settings(**_base(max_ppt_size_mb=2))
    assert s.max_ppt_size_bytes == 2 * 1024 * 1024


def test_download_allowed_hosts_parses_comma_separated():
    s = Settings(**_base(download_allowed_hosts="cdn.example.com, storage.example.com"))
    assert s.download_allowed_hosts == ["cdn.example.com", "storage.example.com"]


def test_production_requires_backend_url():
    with pytest.raises(Exception, match="BACKEND_BASE_URL"):
        Settings(**_base(app_env="production", backend_base_url=""))


def test_production_does_not_read_a_worker_token_from_env():
    s = Settings(**_base(app_env="production", ppt_worker_token=""))
    assert s.is_production is True
    assert "ppt_worker_token" not in s.model_fields


def test_production_allows_local_fs_because_ispring_cloud_holds_the_files():
    s = Settings(**_base(app_env="production", storage_backend="local_fs"))
    assert s.is_production is True


def test_no_apikey_or_database_url_or_invented_routes():
    s = Settings(**_base())
    assert "database_url" not in s.model_fields
    assert "worker_api_key" not in s.model_fields
    assert "backend_api_key" not in s.model_fields
    assert "backend_get_job_path" not in s.model_fields
    assert "backend_callback_path" not in s.model_fields
