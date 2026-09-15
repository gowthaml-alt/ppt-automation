"""Shared synchronous HTTP client. Timeouts are always explicit."""

from __future__ import annotations

import httpx


def create_sync_client(*, timeout_s: float, connect_s: float = 10.0) -> httpx.Client:
    return httpx.Client(
        timeout=httpx.Timeout(timeout_s, connect=min(connect_s, timeout_s)),
        follow_redirects=True,
        headers={"User-Agent": "ppt-automation-worker/1.0"},
    )
