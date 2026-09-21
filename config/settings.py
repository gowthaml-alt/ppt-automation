"""Application configuration loaded from the environment."""

from __future__ import annotations

from functools import lru_cache
from typing import Annotated, Literal

from pydantic import Field, computed_field, field_validator, model_validator
from pydantic_settings import BaseSettings, NoDecode, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,
        extra="ignore",
    )

    app_env: Literal["development", "staging", "production"] = "development"
    log_level: str = "INFO"

    backend_base_url: str = ""

    poll_interval_seconds: int = 30
    get_job_retry_attempts: int = 3
    callback_retry_attempts: int = 3

    temp_root: str = "D:\\ppt-automation\\jobs"
    log_root: str = "D:\\ppt-automation\\logs"
    keep_failed_job_files: bool = False
    min_free_disk_gb: int = 10

    max_ppt_size_mb: int = 500
    download_timeout_seconds: int = 300
    powerpoint_start_timeout_seconds: int = 60
    powerpoint_open_timeout_seconds: int = 120
    ispring_publish_timeout_seconds: int = 1800
    upload_timeout_seconds: int = 600
    backend_request_timeout_seconds: int = 60
    slack_request_timeout_seconds: int = 15
    browser_test_timeout_seconds: int = 60

    slack_webhook_url: str = ""

    powerpoint_exe_path: str = ""
    ispring_install_path: str = ""
    powerpoint_kill_orphans: bool = False
    powerpoint_orphan_max_age_seconds: int = 3600

    # vba and cli were removed: probing the installed Suite 11 showed the
    # add-in has no automation object, no macro entry and no command line.
    ispring_adapter: Literal["not_configured", "fake", "uia"] = "not_configured"
    # Folder in iSpring Cloud that holds one project per institution.
    ispring_parent_folder: str = "PPT Migration"
    # Chrome to attach to for the share step. That browser must already be
    # signed in to iSpring Cloud; see scripts/start_ispring_chrome.cmd.
    ispring_chrome_cdp_url: str = "http://127.0.0.1:9222"
    # Its own profile folder: since Chrome 136 the debug port is ignored on the
    # default profile, so the everyday signed-in Chrome cannot be attached to.
    ispring_chrome_profile_dir: str = r"C:\ispring-chrome-profile"
    ispring_chrome_path: str = ""
    # The iSpring Cloud library. The share step opens this itself: Manage
    # Content opens the machine's default browser, which is not the one the
    # automation attaches to.
    ispring_cloud_url: str = "https://harshit.ispring.com/"
    storage_backend: Literal["local_fs", "s3", "http_api"] = "local_fs"
    browser_test_enabled: bool = True

    download_allowed_hosts: Annotated[list[str], NoDecode] = Field(
        default_factory=list
    )

    s3_bucket: str = ""
    s3_region: str = ""
    s3_endpoint_url: str = ""
    s3_access_key_id: str = ""
    s3_secret_access_key: str = ""
    s3_key_prefix: str = ""
    s3_public_base_url: str = ""

    local_storage_root: str = ""
    local_storage_public_base_url: str = ""

    @field_validator("download_allowed_hosts", mode="before")
    @classmethod
    def _split_hosts(cls, value: object) -> object:
        if isinstance(value, str):
            return [host.strip() for host in value.split(",") if host.strip()]
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
        if not self.backend_base_url:
            raise ValueError(
                "BACKEND_BASE_URL must be set when APP_ENV=production."
            )
        return self


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    return Settings()


def reset_settings_cache() -> None:
    """Clear the cached settings. Tests use this after changing the environment."""
    get_settings.cache_clear()
