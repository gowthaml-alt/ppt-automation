"""Exception hierarchy. Every failure carries the pipeline stage it occurred in."""

from __future__ import annotations


class PptAutomationError(Exception):
    """Base class for every expected failure in this worker."""

    stage: str = "unexpected"

    def __init__(
        self,
        message: str,
        *,
        stage: str | None = None,
        user_message: str | None = None,
    ) -> None:
        super().__init__(message)
        self.message = message
        if stage is not None:
            self.stage = stage
        self.user_message = user_message or message


class JobFetchError(PptAutomationError):
    stage = "fetch"


class LowDiskSpaceError(PptAutomationError):
    stage = "prepare"


class DownloadError(PptAutomationError):
    stage = "download"


class PowerPointAutomationError(PptAutomationError):
    stage = "powerpoint"


class ISpringPublishingError(PptAutomationError):
    stage = "ispring_publish"


class ISpringNotConfiguredError(ISpringPublishingError):
    """Raised by the default adapter: the real mechanism is not yet verified."""


class ISpringTimeoutError(ISpringPublishingError):
    """Publishing exceeded ISPRING_PUBLISH_TIMEOUT_SECONDS."""


class OutputValidationError(PptAutomationError):
    stage = "output_validate"


class OutputUploadError(PptAutomationError):
    stage = "upload"


class BrowserTestError(PptAutomationError):
    stage = "browser_test"


class CallbackError(PptAutomationError):
    stage = "callback"


class CleanupError(PptAutomationError):
    stage = "cleanup"
