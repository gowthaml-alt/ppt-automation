import pytest

from utils.exceptions import DownloadError
from utils.validators import assert_looks_like_pptx, validate_download_url

PPTX_MAGIC = b"PK\x03\x04"


def test_https_url_is_accepted():
    url = "https://storage.example.com/input/biology.pptx"
    assert validate_download_url(url, []) == url


def test_signed_url_is_accepted_unchanged():
    url = "https://s3.example.com/in/a.pptx?X-Amz-Signature=deadbeef"
    assert validate_download_url(url, []) == url


def test_non_http_schemes_are_rejected():
    with pytest.raises(DownloadError, match="scheme"):
        validate_download_url("file:///C:/Windows/win.ini", [])


def test_error_message_does_not_leak_the_query_string():
    signed = "https://evil.example/a.pptx?X-Amz-Signature=deadbeef"
    with pytest.raises(DownloadError) as exc:
        validate_download_url(signed, ["cdn.example.com"])
    assert "deadbeef" not in str(exc.value)


def test_html_error_page_saved_as_pptx_is_rejected(tmp_path):
    path = tmp_path / "a.pptx"
    path.write_bytes(b"<!DOCTYPE html><html><body>403 Forbidden</body></html>")
    with pytest.raises(DownloadError, match="not a valid PowerPoint"):
        assert_looks_like_pptx(path)


def test_valid_pptx_passes(tmp_path):
    path = tmp_path / "a.pptx"
    path.write_bytes(PPTX_MAGIC + b"rest of the zip")
    assert assert_looks_like_pptx(path) is None
