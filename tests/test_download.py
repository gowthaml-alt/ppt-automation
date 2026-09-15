from pathlib import Path

import httpx
import pytest

from config.settings import Settings
from downloader.pptx import download_pptx
from utils.exceptions import DownloadError
from utils.paths import build_job_paths, create_job_dirs

PPTX = b"PK\x03\x04" + b"payload-bytes"


def _settings(tmp_path: Path, **overrides) -> Settings:
    values = {
        "app_env": "development",
        "backend_base_url": "https://backend.example.com",
        "temp_root": str(tmp_path / "jobs"),
        "log_root": str(tmp_path / "logs"),
        "max_ppt_size_mb": 1,
        "download_timeout_seconds": 5,
        "download_allowed_hosts": [],
    }
    values.update(overrides)
    return Settings(**values)


def _paths(tmp_path: Path):
    paths = build_job_paths(tmp_path / "jobs", 101)
    create_job_dirs(paths)
    return paths


def _client(handler) -> httpx.Client:
    return httpx.Client(transport=httpx.MockTransport(handler))


def test_downloads_to_source_pptx_and_removes_part(tmp_path):
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(200, content=PPTX)

    paths = _paths(tmp_path)
    result = download_pptx(
        "https://storage.example.com/input/biology.pptx",
        paths,
        _settings(tmp_path),
        client=_client(handler),
    )
    assert result == paths.source_pptx
    assert result.read_bytes() == PPTX
    assert not paths.partial_pptx.exists()


def test_http_error_is_download_error(tmp_path):
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(404, content=b"missing")

    paths = _paths(tmp_path)
    with pytest.raises(DownloadError):
        download_pptx(
            "https://storage.example.com/missing.pptx",
            paths,
            _settings(tmp_path),
            client=_client(handler),
        )
    assert not paths.source_pptx.exists()


def test_html_error_page_fails_magic_bytes(tmp_path):
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(200, content=b"<html>not a pptx</html>")

    paths = _paths(tmp_path)
    with pytest.raises(DownloadError):
        download_pptx(
            "https://storage.example.com/fake.pptx",
            paths,
            _settings(tmp_path),
            client=_client(handler),
        )
