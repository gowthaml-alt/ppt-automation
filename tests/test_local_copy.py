from pathlib import Path

import pytest

from downloader.local import copy_local_pptx
from tests.pptx_bytes import encrypted_pptx_bytes, minimal_pptx_bytes
from utils.exceptions import DownloadError
from utils.paths import build_job_paths, create_job_dirs

PPTX = minimal_pptx_bytes()


def test_copy_local_pptx_writes_source_pptx(tmp_path):
    src = tmp_path / "sample.pptx"
    src.write_bytes(PPTX)
    paths = build_job_paths(tmp_path / "jobs", 101)
    create_job_dirs(paths)
    dest = copy_local_pptx(src, paths)
    assert dest == paths.source_pptx
    assert dest.read_bytes() == PPTX


def test_copy_local_pptx_rejects_missing_file(tmp_path):
    paths = build_job_paths(tmp_path / "jobs", 101)
    create_job_dirs(paths)
    with pytest.raises(DownloadError):
        copy_local_pptx(tmp_path / "missing.pptx", paths)


def test_copy_local_pptx_rejects_html(tmp_path):
    src = tmp_path / "fake.pptx"
    src.write_bytes(b"<html>nope</html>")
    paths = build_job_paths(tmp_path / "jobs", 101)
    create_job_dirs(paths)
    with pytest.raises(DownloadError):
        copy_local_pptx(src, paths)


def test_copy_local_pptx_rejects_password_protected(tmp_path):
    src = tmp_path / "secret.pptx"
    src.write_bytes(encrypted_pptx_bytes())
    paths = build_job_paths(tmp_path / "jobs", 101)
    create_job_dirs(paths)
    with pytest.raises(DownloadError) as exc:
        copy_local_pptx(src, paths)
    assert exc.value.user_message == "The PowerPoint file is password-protected."
    assert not paths.source_pptx.exists()
