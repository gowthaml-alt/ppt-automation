from utils.exceptions import (
    BrowserTestError,
    CallbackError,
    CleanupError,
    DownloadError,
    ISpringNotConfiguredError,
    ISpringPublishingError,
    ISpringTimeoutError,
    JobFetchError,
    LowDiskSpaceError,
    OutputUploadError,
    OutputValidationError,
    PowerPointAutomationError,
    PptAutomationError,
)

ALL_ERRORS = (
    JobFetchError,
    DownloadError,
    PowerPointAutomationError,
    ISpringPublishingError,
    ISpringNotConfiguredError,
    ISpringTimeoutError,
    OutputValidationError,
    OutputUploadError,
    BrowserTestError,
    CallbackError,
    CleanupError,
    LowDiskSpaceError,
)


def test_every_error_is_a_pptautomationerror():
    for cls in ALL_ERRORS:
        assert issubclass(cls, PptAutomationError)


def test_each_subclass_declares_a_default_stage():
    stages = {
        JobFetchError: "fetch",
        DownloadError: "download",
        PowerPointAutomationError: "powerpoint",
        ISpringPublishingError: "ispring_publish",
        OutputValidationError: "output_validate",
        OutputUploadError: "upload",
        BrowserTestError: "browser_test",
        CallbackError: "callback",
        CleanupError: "cleanup",
        LowDiskSpaceError: "prepare",
    }
    for cls, stage in stages.items():
        assert cls("boom").stage == stage


def test_timeout_is_a_publishing_error():
    assert isinstance(ISpringTimeoutError("timed out"), ISpringPublishingError)


def test_user_message_defaults_to_message():
    assert DownloadError("raw").user_message == "raw"
    err = DownloadError("raw", user_message="Could not download.")
    assert err.user_message == "Could not download."
