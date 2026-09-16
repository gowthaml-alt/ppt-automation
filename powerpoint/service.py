"""Desktop PowerPoint automation. No iSpring logic lives here."""

from __future__ import annotations

import logging
import subprocess
import sys
import threading
from collections.abc import Callable
from pathlib import Path

from config.settings import Settings
from powerpoint import com as windows_com
from utils.exceptions import PowerPointAutomationError
from utils.validators import USER_PASSWORD_PROTECTED

logger = logging.getLogger(__name__)

POWERPOINT_PROCESS = "POWERPNT.EXE"
POWERPOINT_PROGID = "PowerPoint.Application"
PP_ALERTS_NONE = 1

USER_PROTECTED_VIEW = "PowerPoint opened the file in Protected View."
USER_DAMAGED = "The PowerPoint file is damaged and could not be opened."
USER_OPEN_TIMEOUT = "Opening the PowerPoint file timed out."
USER_SAVE_FAILED = "The repaired PowerPoint file could not be saved."


def terminate_powerpoint_processes() -> int:
    """Force-kill desktop PowerPoint. Returns 0 when none is running."""
    if sys.platform != "win32":
        return 0
    completed = subprocess.run(
        ["taskkill", "/F", "/IM", POWERPOINT_PROCESS],
        capture_output=True,
        text=True,
        check=False,
    )
    if completed.returncode in {0, 128}:
        return 0
    logger.warning(
        "taskkill POWERPNT.EXE returned %s: %s",
        completed.returncode,
        (completed.stderr or completed.stdout or "").strip(),
        extra={"stage": "powerpoint_open"},
    )
    return completed.returncode


class PowerPointService:
    def __init__(
        self,
        settings: Settings,
        *,
        terminate_processes: Callable[[], int] | None = None,
    ) -> None:
        self._settings = settings
        self._app = None
        self._presentation = None
        self._terminate_processes = terminate_processes or terminate_powerpoint_processes
        self._open_cleanup_done = False

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
            self._app.DisplayAlerts = PP_ALERTS_NONE
            self._app.Visible = True
        except Exception as exc:
            logger.warning("could not set PowerPoint automation flags", exc_info=True)
            if _is_com(exc):
                raise windows_com.translate_com_error(exc) from exc

    def open(self, path: Path, timeout_s: float | None = None) -> None:
        if self._app is None:
            raise PowerPointAutomationError(
                "open() called before start()",
                user_message="PowerPoint was not started.",
            )
        resolved = Path(path).resolve()
        if not resolved.exists():
            raise PowerPointAutomationError(
                f"presentation does not exist: {resolved.name}",
                user_message=USER_DAMAGED,
            )
        limit = (
            float(self._settings.powerpoint_open_timeout_seconds)
            if timeout_s is None
            else float(timeout_s)
        )
        self._open_cleanup_done = False
        repair_attempted = False
        saved = False
        timed_out = threading.Event()
        done = threading.Event()
        logged_result: dict[str, str | None] = {"open_result": None}

        def log_open(open_result: str, *, did_save: bool = False) -> None:
            if logged_result["open_result"] is not None:
                return
            logged_result["open_result"] = open_result
            logger.info(
                "presentation open finished",
                extra={
                    "stage": "powerpoint_open",
                    "repair_attempted": repair_attempted,
                    "saved": did_save,
                    "open_result": open_result,
                },
            )

        def watchdog() -> None:
            if done.wait(limit):
                return
            timed_out.set()
            log_open("timeout")
            self._cleanup_failed_open()

        open2007 = getattr(getattr(self._app, "Presentations", None), "Open2007", None)
        if open2007 is None:
            log_open("damaged")
            self._cleanup_failed_open()
            raise PowerPointAutomationError(
                "Presentations.Open2007 is not available",
                user_message=USER_DAMAGED,
            )

        repair_attempted = True
        logger.info(
            "opening presentation with OpenAndRepair",
            extra={"stage": "powerpoint_open"},
        )
        watcher = threading.Thread(
            target=watchdog, daemon=True, name="ppt-open-timeout"
        )
        watcher.start()
        try:
            self._presentation = open2007(
                str(resolved),
                False,
                False,
                True,
                True,
            )
            if timed_out.is_set():
                raise PowerPointAutomationError(
                    f"Open2007 exceeded {limit}s",
                    user_message=USER_OPEN_TIMEOUT,
                )
            self._reject_protected_view()
            if self._presentation is not None and not bool(
                getattr(self._presentation, "Saved", True)
            ):
                try:
                    self._presentation.Save()
                    saved = True
                except Exception as exc:
                    log_open("save_failed")
                    self._cleanup_failed_open()
                    raise PowerPointAutomationError(
                        f"Save after repair failed: {exc}",
                        user_message=USER_SAVE_FAILED,
                    ) from exc
            if timed_out.is_set():
                raise PowerPointAutomationError(
                    f"Open2007 exceeded {limit}s",
                    user_message=USER_OPEN_TIMEOUT,
                )
            log_open("success", did_save=saved)
        except PowerPointAutomationError as exc:
            if timed_out.is_set():
                log_open("timeout")
                self._cleanup_failed_open()
                raise PowerPointAutomationError(
                    str(exc),
                    user_message=USER_OPEN_TIMEOUT,
                ) from exc
            if logged_result["open_result"] is None:
                log_open(_result_for_user_message(exc.user_message))
            self._cleanup_failed_open()
            raise
        except Exception as exc:
            if timed_out.is_set():
                log_open("timeout")
                self._cleanup_failed_open()
                raise PowerPointAutomationError(
                    f"Open2007 exceeded {limit}s: {exc}",
                    user_message=USER_OPEN_TIMEOUT,
                ) from exc
            if _is_password_error(exc):
                log_open("password")
                self._cleanup_failed_open()
                raise PowerPointAutomationError(
                    f"password-protected presentation: {exc}",
                    user_message=USER_PASSWORD_PROTECTED,
                ) from exc
            log_open("damaged")
            self._cleanup_failed_open()
            raise PowerPointAutomationError(
                f"Open2007 failed: {exc}",
                user_message=USER_DAMAGED,
            ) from exc
        finally:
            done.set()
            watcher.join(timeout=1)

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

    def _reject_protected_view(self) -> None:
        windows = getattr(self._app, "ProtectedViewWindows", None)
        if windows is None:
            return
        try:
            count = int(windows.Count)
        except Exception:
            return
        if count <= 0:
            return
        for index in range(count, 0, -1):
            try:
                windows.Item(index).Close()
            except Exception:
                logger.warning(
                    "could not close Protected View window", exc_info=True
                )
        raise PowerPointAutomationError(
            "presentation opened in Protected View",
            user_message=USER_PROTECTED_VIEW,
        )

    def _cleanup_failed_open(self) -> None:
        if self._open_cleanup_done:
            return
        self._open_cleanup_done = True
        try:
            self.quit()
        except Exception:
            logger.warning(
                "PowerPoint teardown after open failure failed", exc_info=True
            )
        try:
            self._terminate_processes()
        except Exception:
            logger.warning(
                "PowerPoint terminate after open failure failed", exc_info=True
            )


class NoOpPowerPointService:
    """Used off Windows and in unit tests. Does not launch PowerPoint."""

    def is_available(self) -> tuple[bool, str]:
        return True, "noop"

    def start(self) -> None:
        return None

    def open(self, path: Path, timeout_s: float | None = None) -> None:
        return None

    def close(self) -> None:
        return None

    def quit(self) -> None:
        return None


def _result_for_user_message(user_message: str | None) -> str:
    mapping = {
        USER_OPEN_TIMEOUT: "timeout",
        USER_PASSWORD_PROTECTED: "password",
        USER_PROTECTED_VIEW: "protected_view",
        USER_SAVE_FAILED: "save_failed",
    }
    return mapping.get(user_message or "", "damaged")


def _is_password_error(exc: BaseException) -> bool:
    text = str(exc).lower()
    if "password" in text:
        return True
    wrapped = windows_com.translate_com_error(exc)
    return "password" in wrapped.message.lower()


def _is_com(exc: BaseException) -> bool:
    return exc.__class__.__module__.startswith("pywintypes") or (
        exc.__class__.__name__ == "com_error"
    )
