from pathlib import Path

import pytest

from config.settings import Settings
from powerpoint.com import translate_com_error
from powerpoint.service import PowerPointService
from utils.exceptions import PowerPointAutomationError


def _settings(tmp_path: Path) -> Settings:
    return Settings(
        app_env="development",
        backend_base_url="https://backend.example.com",
        temp_root=str(tmp_path / "jobs"),
        log_root=str(tmp_path / "logs"),
    )


def test_translate_com_error_extracts_hresult():
    class FakeCom(Exception):
        pass

    exc = FakeCom((0x80040154, "Class not registered"))
    wrapped = translate_com_error(exc)
    assert isinstance(wrapped, PowerPointAutomationError)
    assert "0x80040154" in wrapped.message or "Class not registered" in wrapped.message


def test_is_available_is_false_off_windows(tmp_path, monkeypatch):
    monkeypatch.setattr("powerpoint.service.sys.platform", "darwin")
    ok, detail = PowerPointService(_settings(tmp_path)).is_available()
    assert ok is False
    assert "Windows" in detail


def test_start_raises_off_windows(tmp_path, monkeypatch):
    monkeypatch.setattr("powerpoint.service.sys.platform", "darwin")
    with pytest.raises(PowerPointAutomationError):
        PowerPointService(_settings(tmp_path)).start()


def test_close_and_quit_are_idempotent_when_never_started(tmp_path):
    service = PowerPointService(_settings(tmp_path))
    service.close()
    service.quit()
