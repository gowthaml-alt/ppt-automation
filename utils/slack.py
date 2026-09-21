"""Optional Slack failure notification. Never raises. Never logs the webhook."""

from __future__ import annotations

import logging

import httpx

from api_client.jobs import Job
from config.settings import Settings
from utils.http import create_sync_client

logger = logging.getLogger(__name__)


class SlackNotifier:
    def __init__(
        self, settings: Settings, *, client: httpx.Client | None = None
    ) -> None:
        self._settings = settings
        self._client = client

    def notify_failure(self, job: Job, *, stage: str, error_message: str) -> None:
        webhook = self._settings.slack_webhook_url
        if not webhook:
            return
        text = (
            f"PPT automation failed\n"
            f"job_id={job.job_id}\n"
            f"material_id={job.material_id}\n"
            f"material_name={job.material_name}\n"
            f"institution_name={job.institution_name or ''}\n"
            f"stage={stage}\n"
            f"error={error_message}"
        )
        owns = self._client is None
        client = self._client or create_sync_client(
            timeout_s=self._settings.slack_request_timeout_seconds
        )
        try:
            client.post(webhook, json={"text": text})
        except Exception:
            logger.warning("slack notification failed", extra={"stage": "slack"}, exc_info=True)
        finally:
            if owns:
                client.close()
