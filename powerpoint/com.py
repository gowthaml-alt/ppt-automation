"""The only module allowed to import pywin32.

Imports are guarded so this file is importable on macOS and Linux.
"""

from __future__ import annotations

import logging
import sys
from collections.abc import Iterator
from contextlib import contextmanager

from utils.exceptions import PowerPointAutomationError

logger = logging.getLogger(__name__)


def is_windows() -> bool:
    return sys.platform == "win32"


def translate_com_error(exc: BaseException) -> PowerPointAutomationError:
    hresult = None
    description = str(exc)
    args = getattr(exc, "args", ())
    if args:
        first = args[0]
        if isinstance(first, tuple) and first:
            hresult = first[0]
            if len(first) > 1:
                description = str(first[1])
        else:
            description = str(first)
    hex_code = (
        f"0x{int(hresult) & 0xFFFFFFFF:08X}" if isinstance(hresult, int) else "unknown"
    )
    return PowerPointAutomationError(
        f"COM error {hex_code}: {description}",
        user_message="PowerPoint automation failed.",
    )


@contextmanager
def com_apartment() -> Iterator[None]:
    if not is_windows():
        yield
        return
    try:
        import pythoncom  # type: ignore
    except ImportError as exc:
        raise PowerPointAutomationError(
            f"pywin32 is not installed: {exc}",
            user_message="PowerPoint automation libraries are not installed.",
        ) from exc
    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()


def dispatch(prog_id: str):
    if not is_windows():
        raise PowerPointAutomationError(
            f"COM dispatch of {prog_id!r} is only available on Windows",
            user_message="PowerPoint is only available on the Windows worker.",
        )
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise PowerPointAutomationError(
            f"pywin32 is not installed: {exc}",
            user_message="PowerPoint automation libraries are not installed.",
        ) from exc
    try:
        return win32com.client.Dispatch(prog_id)
    except Exception as exc:
        logger.error("COM dispatch failed", extra={"prog_id": prog_id}, exc_info=True)
        raise translate_com_error(exc) from exc
