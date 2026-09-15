from pathlib import Path

import pytest

from config.settings import Settings
from storage.backend import get_storage_backend, mime_type_for, output_prefix
from utils.exceptions import OutputUploadError
from validator.html5 import validate_ispring_output


def _package(root: Path) -> Path:
    (root / "data").mkdir(parents=True)
    (root / "index.html").write_text("<html>ok</html>", encoding="utf-8")
    (root / "data" / "app.js").write_text("console.log(1)", encoding="utf-8")
    return root


def _settings(tmp_path: Path, **overrides) -> Settings:
    values = {
        "app_env": "development",
        "backend_base_url": "https://backend.example.com",
        "temp_root": str(tmp_path / "jobs"),
        "log_root": str(tmp_path / "logs"),
        "storage_backend": "local_fs",
        "local_storage_root": str(tmp_path / "cdn"),
        "local_storage_public_base_url": "https://cdn.example.com",
    }
    values.update(overrides)
    return Settings(**values)


def test_mime_types_are_explicit():
    assert mime_type_for(Path("x.js")) == "application/javascript"
    assert mime_type_for(Path("x.wasm")) == "application/wasm"


def test_output_prefix_matches_iframe_example():
    settings = Settings(
        backend_base_url="https://backend.example.com",
        temp_root="/tmp/j",
        log_root="/tmp/l",
    )
    assert output_prefix(5001, settings) == "ppt/5001"


def test_local_fs_uploads_tree_and_builds_iframe_url(tmp_path):
    backend = get_storage_backend(_settings(tmp_path))
    result = backend.upload_directory(_package(tmp_path / "output"), prefix="ppt/5001")
    assert result.entry_file == "index.html"
    assert result.iframe_url == "https://cdn.example.com/ppt/5001/index.html"
    validate_ispring_output(tmp_path / "cdn" / "ppt" / "5001")


def test_http_api_backend_is_a_documented_stub(tmp_path):
    backend = get_storage_backend(_settings(tmp_path, storage_backend="http_api"))
    with pytest.raises(OutputUploadError):
        backend.upload_directory(_package(tmp_path / "output"), prefix="x")
