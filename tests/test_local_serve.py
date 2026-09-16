from pathlib import Path

import pytest

from browser_test.local_serve import LocalOutputChecker, http_get_check
from config.settings import Settings
from utils.exceptions import BrowserTestError


def _package(root: Path) -> Path:
    (root / "data").mkdir(parents=True)
    (root / "index.html").write_text(
        "<html><head><title>ok</title></head><body>published</body></html>",
        encoding="utf-8",
    )
    (root / "data" / "app.js").write_text("window.OK=true;", encoding="utf-8")
    return root


def test_http_get_check_passes_for_html(tmp_path):
    package = _package(tmp_path / "ppt" / "5001")
    checker = LocalOutputChecker(timeout_s=5)
    url = checker.serve_and_check(package)
    assert url.startswith("http://127.0.0.1:")
    assert url.endswith("/index.html")


def test_http_get_check_fails_when_index_missing(tmp_path):
    empty = tmp_path / "empty"
    empty.mkdir()
    with pytest.raises(BrowserTestError):
        LocalOutputChecker(timeout_s=5).serve_and_check(empty)


def test_http_get_check_rejects_unreachable_url():
    with pytest.raises(BrowserTestError):
        http_get_check("http://127.0.0.1:1/", timeout_s=1)
