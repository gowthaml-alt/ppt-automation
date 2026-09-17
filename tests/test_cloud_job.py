import zipfile
from pathlib import Path

import pytest

from config.settings import Settings
from tests.pptx_bytes import minimal_pptx_bytes
from utils.exceptions import DownloadError
from worker.cloud_job import filename_from_url, inspect

SIGNED_URL = (
    "https://dme2wmiz2suov.cloudfront.net/User(93283564)/Course(134603)/"
    "Section(536002)/20242897-Kickoff_and_Advanced_Prompting.pptx"
    "?Expires=1821085378&Signature=so8FOe7m4JDSfR8Gc~AnLxr9zJ--&"
    "Key-Pair-Id=APKAIIFZDCEANAVU2VTA"
)


def _settings(tmp_path: Path, **overrides) -> Settings:
    values = {
        "app_env": "development",
        "backend_base_url": "https://backend.example.com",
        "temp_root": str(tmp_path / "jobs"),
        "log_root": str(tmp_path / "logs"),
    }
    values.update(overrides)
    return Settings(**values)


def test_filename_ignores_the_signature_query():
    assert (
        filename_from_url(SIGNED_URL)
        == "20242897-Kickoff_and_Advanced_Prompting.pptx"
    )


def test_filename_handles_escapes_and_empty_paths():
    assert filename_from_url("https://x.test/a/My%20Deck.pptx") == "My Deck.pptx"
    assert filename_from_url("https://x.test/") == "download.pptx"
    assert filename_from_url("") == "download.pptx"


def test_good_package_is_not_flagged_as_damaged(tmp_path):
    deck = tmp_path / "good.pptx"
    deck.write_bytes(minimal_pptx_bytes())
    assert inspect(deck) is False


def test_damaged_zip_is_passed_to_powerpoint_not_rejected(tmp_path):
    """A broken package must reach PowerPoint: repair is the only thing that fixes it."""
    deck = tmp_path / "broken.pptx"
    deck.write_bytes(b"PK\x03\x04" + b"this is not really a zip" * 10)
    assert inspect(deck) is True


def test_missing_parts_are_passed_to_powerpoint(tmp_path):
    deck = tmp_path / "incomplete.pptx"
    with zipfile.ZipFile(deck, "w") as archive:
        archive.writestr("docProps/app.xml", "<Properties/>")
    assert inspect(deck) is True


def test_password_protected_fails_before_powerpoint(tmp_path):
    """Repair cannot undo a password, so this must not waste a PowerPoint run."""
    deck = tmp_path / "locked.pptx"
    with zipfile.ZipFile(deck, "w") as archive:
        archive.writestr("EncryptionInfo", "x")
        archive.writestr("EncryptedPackage", "x")
    with pytest.raises(DownloadError) as exc:
        inspect(deck)
    assert "password" in (exc.value.user_message or "").lower()


def test_a_file_that_is_not_a_presentation_fails(tmp_path):
    deck = tmp_path / "notppt.pptx"
    deck.write_bytes(b"<html>login page</html>")
    with pytest.raises(DownloadError):
        inspect(deck)


def test_old_binary_ppt_is_accepted(tmp_path):
    """PowerPoint opens legacy .ppt; the zip checks do not apply to it."""
    deck = tmp_path / "legacy.ppt"
    deck.write_bytes(b"\xd0\xcf\x11\xe0" + b"\x00" * 64)
    assert inspect(deck) is False


def test_download_rejects_a_disallowed_host(tmp_path):
    from worker.cloud_job import download

    settings = _settings(tmp_path, download_allowed_hosts=["only.example.com"])
    with pytest.raises(DownloadError):
        download(SIGNED_URL, tmp_path, settings)
