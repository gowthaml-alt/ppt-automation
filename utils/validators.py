"""Input validation for untrusted values."""

from __future__ import annotations

import zipfile
from pathlib import Path
from urllib.parse import urlsplit

from utils.exceptions import DownloadError
from utils.logging_config import scrub_url

ALLOWED_SCHEMES = frozenset({"http", "https"})
OOXML_MAGIC = b"PK\x03\x04"
REQUIRED_PPTX_PARTS = ("[Content_Types].xml", "ppt/presentation.xml")
ENCRYPTED_PPTX_NAMES = frozenset({"EncryptionInfo", "EncryptedPackage"})

USER_NOT_POWERPOINT = "The downloaded file is not a valid PowerPoint file."
USER_INVALID_PACKAGE = "The PowerPoint file is not a valid package."
USER_MISSING_PARTS = "The PowerPoint file is missing required package parts."
USER_PASSWORD_PROTECTED = "The PowerPoint file is password-protected."


def validate_download_url(url: str, allowed_hosts: list[str]) -> str:
    """Validate an input PPT URL.

    Restricts the scheme to http/https and, when ``allowed_hosts`` is
    non-empty, restricts the host to an exact match. An empty allowlist
    permits any host; the resulting SSRF exposure is documented in the README.
    """
    parts = urlsplit(url or "")
    if parts.scheme.lower() not in ALLOWED_SCHEMES:
        raise DownloadError(
            f"unsupported URL scheme {parts.scheme!r}; only http and https are allowed",
            user_message="The PPT URL scheme is not supported.",
        )
    if not parts.hostname:
        raise DownloadError(
            "URL has no host",
            user_message="The PPT URL has no host.",
        )
    if allowed_hosts:
        permitted = {host.lower() for host in allowed_hosts}
        if parts.hostname.lower() not in permitted:
            raise DownloadError(
                f"host {parts.hostname!r} is not allowed by DOWNLOAD_ALLOWED_HOSTS "
                f"(url: {scrub_url(url)})",
                user_message="The PPT URL host is not permitted.",
            )
    return url


def assert_looks_like_pptx(path: Path) -> None:
    """Raise unless ``path`` is a non-empty file starting with ZIP magic bytes."""
    if not path.exists():
        raise DownloadError(
            f"downloaded file does not exist: {path.name}",
            user_message="The downloaded file is missing.",
        )
    if path.stat().st_size == 0:
        raise DownloadError(
            f"downloaded file is empty: {path.name}",
            user_message="The downloaded file is empty.",
        )
    with path.open("rb") as handle:
        header = handle.read(len(OOXML_MAGIC))
    if header != OOXML_MAGIC:
        raise DownloadError(
            f"downloaded file is not a valid PowerPoint package "
            f"(header {header!r}, expected {OOXML_MAGIC!r})",
            user_message=USER_NOT_POWERPOINT,
        )


def assert_valid_pptx_package(path: Path) -> None:
    """Raise unless ``path`` is a readable PPTX with required OPC parts."""
    try:
        with zipfile.ZipFile(path) as archive:
            names = {name.replace("\\", "/") for name in archive.namelist()}
    except zipfile.BadZipFile as exc:
        raise DownloadError(
            f"downloaded file is not a valid PowerPoint package: {path.name}",
            user_message=USER_INVALID_PACKAGE,
        ) from exc
    except OSError as exc:
        raise DownloadError(
            f"could not read downloaded file {path.name}: {exc}",
            user_message=USER_INVALID_PACKAGE,
        ) from exc

    basenames = {name.rsplit("/", 1)[-1] for name in names}
    if ENCRYPTED_PPTX_NAMES & basenames:
        raise DownloadError(
            f"downloaded file is password-protected: {path.name}",
            user_message=USER_PASSWORD_PROTECTED,
        )

    missing = [part for part in REQUIRED_PPTX_PARTS if part not in names]
    if missing:
        raise DownloadError(
            f"downloaded file is missing package parts {missing}: {path.name}",
            user_message=USER_MISSING_PARTS,
        )
