"""Desktop PowerPoint automation. No iSpring logic lives here."""

from __future__ import annotations

import logging
import sys
from pathlib import Path

from config.settings import Settings
from powerpoint import com as windows_com
from utils.exceptions import PowerPointAutomationError

logger = logging.getLogger(__name__)

POWERPOINT_PROCESS = "POWERPNT.EXE"
POWERPOINT_PROGID = "PowerPoint.Application"


class PowerPointService:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        self._app = None
        self._presentation = None

    def is_available(self) -> tuple[bool, str]:
        if sys.platform != "win32":
            return False, "PowerPoint COM is only available on Windows"
        configured = self._settings.powerpoint_exe_path
        if configured and not Path(configured).exists():
            return False, "POWERPOINT_EXE_PATH does not exist"
        return True, "Windows host; COM will be opened at job time"

    def start(self) -> None:
        if sys.platform != "win32":
            raise PowerPointAutomationError(
                "PowerPoint cannot be started off Windows",
                user_message="PowerPoint is only available on the Windows worker.",
            )
        logger.info("starting PowerPoint", extra={"stage": "powerpoint_open"})
        self._app = windows_com.dispatch(POWERPOINT_PROGID)
        try:
            self._app.DisplayAlerts = 1
            self._app.Visible = True
        except Exception as exc:
            logger.warning("could not set PowerPoint automation flags", exc_info=True)
            if _is_com(exc):
                raise windows_com.translate_com_error(exc) from exc

    def open(self, path: Path) -> None:
        if self._app is None:
            raise PowerPointAutomationError(
                "open() called before start()",
                user_message="PowerPoint was not started.",
            )
        resolved = Path(path).resolve()
        if not resolved.exists():
            raise PowerPointAutomationError(
                f"presentation does not exist: {resolved.name}",
                user_message="The PowerPoint file could not be opened.",
            )
        logger.info("opening presentation", extra={"stage": "powerpoint_open"})
        try:
            self._presentation = self._app.Presentations.Open(
                str(resolved), False, False, True
            )
        except Exception as exc:
            logger.error("Presentations.Open failed", exc_info=True)
            raise windows_com.translate_com_error(exc) from exc

    def close(self) -> None:
        presentation = self._presentation
        self._presentation = None
        if presentation is None:
            return
        try:
            presentation.Close()
        except Exception:
            logger.warning("presentation.Close failed", exc_info=True)

    def quit(self) -> None:
        self.close()
        app = self._app
        self._app = None
        if app is None:
            return
        try:
            app.Quit()
        except Exception:
            logger.warning("PowerPoint Quit failed", exc_info=True)


class NoOpPowerPointService:
    """Used off Windows and in unit tests. Does not launch PowerPoint."""

    def is_available(self) -> tuple[bool, str]:
        return True, "noop"

    def start(self) -> None:
        return None

    def open(self, path: Path) -> None:
        return None

    def close(self) -> None:
        return None

    def quit(self) -> None:
        return None


def _is_com(exc: BaseException) -> bool:
    return exc.__class__.__module__.startswith("pywintypes") or (
        exc.__class__.__name__ == "com_error"
    )
