"""Serve uploaded HTML5 locally and verify index.html loads.

Used when STORAGE_BACKEND=local_fs so Playwright/httpx can hit a real HTTP URL
instead of a dummy public CDN host.
"""

from __future__ import annotations

import logging
import threading
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import httpx

from utils.exceptions import BrowserTestError
from utils.logging_config import scrub_url

logger = logging.getLogger(__name__)


def http_get_check(url: str, *, timeout_s: float) -> None:
    try:
        response = httpx.get(url, timeout=timeout_s, follow_redirects=True)
    except httpx.HTTPError as exc:
        raise BrowserTestError(
            f"HTTP check failed for {scrub_url(url)}: {exc}",
            user_message="The published page could not be fetched.",
        ) from exc
    if response.status_code >= 400:
        raise BrowserTestError(
            f"HTTP check returned {response.status_code} for {scrub_url(url)}",
            user_message="The published page returned an error status.",
        )
    body = response.text
    if "<html" not in body.lower() or len(body.strip()) < 32:
        raise BrowserTestError(
            "published page is empty or not HTML",
            user_message="The published page does not look like HTML5 output.",
        )


class LocalOutputChecker:
    def __init__(self, *, timeout_s: float = 15, use_playwright: bool = False) -> None:
        self._timeout_s = timeout_s
        self._use_playwright = use_playwright

    def serve_and_check(self, package_dir: Path) -> str:
        package_dir = Path(package_dir).resolve()
        index = package_dir / "index.html"
        if not index.is_file() or index.stat().st_size == 0:
            raise BrowserTestError(
                f"uploaded output is missing index.html: {package_dir}",
                user_message="Upload succeeded but index.html is missing.",
            )

        handler = partial(SimpleHTTPRequestHandler, directory=str(package_dir))
        server = ThreadingHTTPServer(("127.0.0.1", 0), handler)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        port = server.server_address[1]
        url = f"http://127.0.0.1:{port}/index.html"
        try:
            logger.info("local HTTP check starting", extra={"stage": "browser_test", "url": url})
            http_get_check(url, timeout_s=self._timeout_s)
            if self._use_playwright:
                _playwright_check(url, timeout_s=self._timeout_s)
            logger.info("local HTTP check passed", extra={"stage": "browser_test"})
            return url
        finally:
            server.shutdown()
            server.server_close()
            thread.join(timeout=2)


def _playwright_check(url: str, *, timeout_s: float) -> None:
    try:
        from playwright.sync_api import Error as PlaywrightError
        from playwright.sync_api import sync_playwright
    except ImportError:
        logger.warning(
            "playwright is not installed; HTTP check already passed",
            extra={"stage": "browser_test"},
        )
        return
    timeout_ms = timeout_s * 1000
    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=True)
            try:
                page = browser.new_page()
                response = page.goto(
                    url, wait_until="domcontentloaded", timeout=timeout_ms
                )
                if response is None or response.status >= 400:
                    status = None if response is None else response.status
                    raise BrowserTestError(
                        f"playwright got HTTP {status} for {url}",
                        user_message="The published HTML5 page failed in the browser.",
                    )
                html = page.content()
                if "<html" not in html.lower():
                    raise BrowserTestError(
                        "playwright loaded a non-HTML document",
                        user_message="The published HTML5 page failed in the browser.",
                    )
            finally:
                browser.close()
    except BrowserTestError:
        raise
    except PlaywrightError as exc:
        raise BrowserTestError(
            f"playwright failed: {exc}",
            user_message="The published HTML5 page failed in the browser.",
        ) from exc
