"""Playwright check of the published HTML5 package in a real browser.

This module does not control the PowerPoint desktop application.
"""

from __future__ import annotations

import logging

from config.settings import Settings
from utils.exceptions import BrowserTestError
from utils.logging_config import scrub_url

logger = logging.getLogger(__name__)


class PlaywrightChecker:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings

    def check(self, iframe_url: str) -> None:
        if not self._settings.browser_test_enabled:
            logger.info(
                "browser test skipped by configuration",
                extra={"stage": "browser_test", "url": scrub_url(iframe_url)},
            )
            return
        try:
            from playwright.sync_api import Error as PlaywrightError
            from playwright.sync_api import sync_playwright
        except ImportError as exc:
            raise BrowserTestError(
                f"playwright is not installed: {exc}",
                user_message="Playwright is not installed on this worker.",
            ) from exc

        timeout_ms = self._settings.browser_test_timeout_seconds * 1000
        logger.info(
            "browser test starting",
            extra={"stage": "browser_test", "url": scrub_url(iframe_url)},
        )
        try:
            with sync_playwright() as playwright:
                browser = playwright.chromium.launch(headless=True)
                try:
                    page = browser.new_page()
                    response = page.goto(
                        iframe_url, wait_until="domcontentloaded", timeout=timeout_ms
                    )
                    if response is None or response.status >= 400:
                        status = None if response is None else response.status
                        raise BrowserTestError(
                            f"iframe URL returned HTTP {status}",
                            user_message="The published HTML5 page could not be loaded.",
                        )
                    html = page.content()
                    if "<html" not in html.lower() or len(html.strip()) < 32:
                        raise BrowserTestError(
                            "iframe document is empty or not HTML",
                            user_message="The published HTML5 page looks empty.",
                        )
                finally:
                    browser.close()
        except BrowserTestError:
            raise
        except PlaywrightError as exc:
            raise BrowserTestError(
                f"playwright failed for {scrub_url(iframe_url)}: {exc}",
                user_message="The published HTML5 page failed browser verification.",
            ) from exc
        logger.info("browser test passed", extra={"stage": "browser_test"})


class NoOpBrowserChecker:
    def check(self, iframe_url: str) -> None:
        logger.info(
            "browser test noop",
            extra={"stage": "browser_test", "url": scrub_url(iframe_url)},
        )
