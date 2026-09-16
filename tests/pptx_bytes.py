"""Minimal PPTX-like zip payloads for unit tests."""

from __future__ import annotations

import io
import zipfile


def minimal_pptx_bytes() -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org'
            '/package/2006/content-types"></Types>',
        )
        archive.writestr(
            "ppt/presentation.xml",
            '<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.'
            'openxmlformats.org/presentationml/2006/main"></p:presentation>',
        )
    return buffer.getvalue()


def encrypted_pptx_bytes() -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("EncryptionInfo", b"encrypted-info")
        archive.writestr("EncryptedPackage", b"encrypted-package")
    return buffer.getvalue()


def zip_missing_presentation_bytes() -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
    return buffer.getvalue()


def invalid_zip_with_magic_bytes() -> bytes:
    return b"PK\x03\x04this-is-not-a-zip-payload"
