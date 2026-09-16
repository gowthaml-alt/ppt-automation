from pathlib import Path
import logging
import time

import pytest

from config.settings import Settings
from powerpoint.com import translate_com_error
from powerpoint.service import PowerPointService
from tests.pptx_bytes import minimal_pptx_bytes
from utils.exceptions import PowerPointAutomationError


def _settings(tmp_path: Path, **overrides) -> Settings:
    values = {
        "app_env": "development",
        "backend_base_url": "https://backend.example.com",
        "temp_root": str(tmp_path / "jobs"),
        "log_root": str(tmp_path / "logs"),
        "powerpoint_open_timeout_seconds": 2,
    }
    values.update(overrides)
    return Settings(**values)


def _working_pptx(tmp_path: Path) -> Path:
    path = tmp_path / "working" / "source.pptx"
    path.parent.mkdir(parents=True)
    path.write_bytes(minimal_pptx_bytes())
    return path


class FakeProtectedViewWindows:
    def __init__(self, count: int = 0) -> None:
        self.Count = count
        self.closed = 0

    def Item(self, index: int):
        parent = self

        class _Window:
            def Close(self) -> None:
                parent.closed += 1
                parent.Count = max(0, parent.Count - 1)

        return _Window()


class FakePresentation:
    def __init__(
        self,
        path: Path,
        *,
        saved: bool = True,
        save_error: Exception | None = None,
        save_as_error: Exception | None = None,
        read_only: bool = False,
    ) -> None:
        self.FullName = str(path)
        self.Saved = saved
        self.ReadOnly = read_only
        self.save_error = save_error
        self.save_as_error = save_as_error
        self.save_calls = 0
        self.save_as_calls: list[tuple[str, int]] = []
        self.closed = 0

    def Save(self) -> None:
        self.save_calls += 1
        if self.save_error is not None:
            raise self.save_error
        Path(self.FullName).write_bytes(b"repaired-working-copy")
        self.Saved = True

    def SaveAs(self, FileName, FileFormat=24):
        self.save_as_calls.append((str(FileName), int(FileFormat)))
        if self.save_as_error is not None:
            raise self.save_as_error
        Path(FileName).write_bytes(b"repaired-working-copy")
        self.FullName = str(FileName)
        self.Saved = True
        self.ReadOnly = False

    def Close(self) -> None:
        self.closed += 1


class FakePresentations:
    def __init__(
        self,
        presentation: FakePresentation | None = None,
        *,
        error: Exception | None = None,
        delay_s: float = 0,
        missing_open2007: bool = False,
    ) -> None:
        self.presentation = presentation
        self.error = error
        self.delay_s = delay_s
        self.missing_open2007 = missing_open2007
        self.calls: list[dict] = []

    def Open2007(self, FileName, ReadOnly=False, Untitled=False, WithWindow=True, OpenAndRepair=False):
        if self.delay_s:
            time.sleep(self.delay_s)
        self.calls.append(
            {
                "FileName": FileName,
                "ReadOnly": ReadOnly,
                "Untitled": Untitled,
                "WithWindow": WithWindow,
                "OpenAndRepair": OpenAndRepair,
            }
        )
        if self.error is not None:
            raise self.error
        return self.presentation

    def Open(self, *args, **kwargs):
        raise AssertionError("Presentations.Open must not be used; use Open2007")


class FakeApp:
    def __init__(self, presentations: FakePresentations, *, protected_count: int = 0) -> None:
        self.Presentations = presentations
        if presentations.missing_open2007:
            del self.Presentations.Open2007
        self.ProtectedViewWindows = FakeProtectedViewWindows(protected_count)
        self.DisplayAlerts = None
        self.Visible = None
        self.quit_calls = 0

    def Quit(self) -> None:
        self.quit_calls += 1


def _service(tmp_path: Path, app: FakeApp, killed: list[int] | None = None) -> PowerPointService:
    service = PowerPointService(
        _settings(tmp_path),
        terminate_processes=lambda: killed.append(1) or len(killed),
    )
    service._app = app
    return service


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


def test_open_uses_open2007_with_repair(tmp_path, caplog):
    caplog.set_level(logging.INFO)
    working = _working_pptx(tmp_path)
    presentation = FakePresentation(working, saved=True)
    presentations = FakePresentations(presentation)
    app = FakeApp(presentations)
    killed: list[int] = []
    _service(tmp_path, app, killed).open(working)
    assert presentations.calls[0]["OpenAndRepair"] is True
    assert presentations.calls[0]["WithWindow"] is True
    assert presentations.calls[0]["ReadOnly"] is False
    assert killed == []
    records = [r for r in caplog.records if getattr(r, "open_result", None)]
    assert records[-1].repair_attempted is True
    assert records[-1].saved is False
    assert records[-1].open_result == "success"


def test_repairable_file_is_saved_to_working_copy_only(tmp_path, caplog):
    caplog.set_level(logging.INFO)
    original = tmp_path / "input" / "source.pptx"
    original.parent.mkdir(parents=True)
    original.write_bytes(minimal_pptx_bytes())
    working = _working_pptx(tmp_path)
    original_bytes = original.read_bytes()
    presentation = FakePresentation(working, saved=False)
    app = FakeApp(FakePresentations(presentation))
    _service(tmp_path, app).open(working)
    assert presentation.save_calls == 1
    assert working.read_bytes() == b"repaired-working-copy"
    assert original.read_bytes() == original_bytes
    records = [r for r in caplog.records if getattr(r, "open_result", None)]
    assert records[-1].saved is True
    assert records[-1].open_result == "success"


def test_readonly_repair_is_saved_under_a_new_name_then_swapped(tmp_path, caplog):
    caplog.set_level(logging.INFO)
    original = tmp_path / "input" / "source.pptx"
    original.parent.mkdir(parents=True)
    original.write_bytes(minimal_pptx_bytes())
    working = _working_pptx(tmp_path)
    original_bytes = original.read_bytes()
    presentation = FakePresentation(
        working,
        saved=False,
        save_error=RuntimeError(
            "Presentation.Save : This presentation is read-only and must be "
            "saved with a different name."
        ),
    )
    presentations = FakePresentations(presentation)
    killed: list[int] = []
    _service(tmp_path, FakeApp(presentations), killed).open(working)
    repaired = working.resolve().with_name("source.repaired.pptx")
    assert presentation.save_calls == 1
    assert presentation.save_as_calls == [(str(repaired), 24)]
    assert presentation.closed == 1
    assert not repaired.exists()
    assert working.read_bytes() == b"repaired-working-copy"
    assert original.read_bytes() == original_bytes
    assert [call["OpenAndRepair"] for call in presentations.calls] == [True, False]
    assert presentations.calls[-1]["FileName"] == str(working.resolve())
    assert killed == []
    records = [r for r in caplog.records if getattr(r, "open_result", None)]
    assert records[-1].saved is True
    assert records[-1].open_result == "success"


def test_read_only_presentation_skips_save_and_goes_straight_to_save_as(tmp_path):
    working = _working_pptx(tmp_path)
    presentation = FakePresentation(working, saved=False, read_only=True)
    presentations = FakePresentations(presentation)
    _service(tmp_path, FakeApp(presentations)).open(working)
    repaired = working.resolve().with_name("source.repaired.pptx")
    assert presentation.save_calls == 0
    assert presentation.save_as_calls == [(str(repaired), 24)]
    assert working.read_bytes() == b"repaired-working-copy"
    assert not repaired.exists()


def test_save_as_failure_after_readonly_save_terminates_powerpoint(tmp_path, caplog):
    caplog.set_level(logging.INFO)
    working = _working_pptx(tmp_path)
    presentation = FakePresentation(
        working,
        saved=False,
        read_only=True,
        save_as_error=RuntimeError("disk full"),
    )
    app = FakeApp(FakePresentations(presentation))
    killed: list[int] = []
    with pytest.raises(PowerPointAutomationError) as exc:
        _service(tmp_path, app, killed).open(working)
    assert (
        exc.value.user_message == "The repaired PowerPoint file could not be saved."
    )
    assert killed == [1]
    records = [r for r in caplog.records if getattr(r, "open_result", None)]
    assert records[-1].open_result == "save_failed"


def test_unrepairable_file_fails_and_terminates_powerpoint(tmp_path, caplog):
    caplog.set_level(logging.INFO)
    working = _working_pptx(tmp_path)
    presentations = FakePresentations(
        error=RuntimeError("COM error 0x80004005: cannot open")
    )
    app = FakeApp(presentations)
    killed: list[int] = []
    with pytest.raises(PowerPointAutomationError) as exc:
        _service(tmp_path, app, killed).open(working)
    assert (
        exc.value.user_message
        == "The PowerPoint file is damaged and could not be opened."
    )
    assert killed == [1]
    assert app.quit_calls == 1
    records = [r for r in caplog.records if getattr(r, "open_result", None)]
    assert records[-1].open_result == "damaged"
    assert records[-1].repair_attempted is True


def test_password_com_error_fails_without_continuing(tmp_path, caplog):
    caplog.set_level(logging.INFO)
    working = _working_pptx(tmp_path)
    presentations = FakePresentations(
        error=RuntimeError("The password you supplied is not correct")
    )
    app = FakeApp(presentations)
    killed: list[int] = []
    with pytest.raises(PowerPointAutomationError) as exc:
        _service(tmp_path, app, killed).open(working)
    assert exc.value.user_message == "The PowerPoint file is password-protected."
    assert killed == [1]
    records = [r for r in caplog.records if getattr(r, "open_result", None)]
    assert records[-1].open_result == "password"


def test_protected_view_fails_without_enabling_editing(tmp_path, caplog):
    caplog.set_level(logging.INFO)
    working = _working_pptx(tmp_path)
    presentation = FakePresentation(working, saved=True)
    app = FakeApp(FakePresentations(presentation), protected_count=1)
    killed: list[int] = []
    with pytest.raises(PowerPointAutomationError) as exc:
        _service(tmp_path, app, killed).open(working)
    assert exc.value.user_message == "PowerPoint opened the file in Protected View."
    assert app.ProtectedViewWindows.closed == 1
    assert presentation.save_calls == 0
    assert killed == [1]
    records = [r for r in caplog.records if getattr(r, "open_result", None)]
    assert records[-1].open_result == "protected_view"


def test_open_timeout_terminates_powerpoint(tmp_path, caplog):
    caplog.set_level(logging.INFO)
    working = _working_pptx(tmp_path)
    presentation = FakePresentation(working, saved=True)
    presentations = FakePresentations(presentation, delay_s=0.4)
    app = FakeApp(presentations)
    killed: list[int] = []
    service = _service(tmp_path, app, killed)
    with pytest.raises(PowerPointAutomationError) as exc:
        service.open(working, timeout_s=0.05)
    assert exc.value.user_message == "Opening the PowerPoint file timed out."
    assert killed == [1]
    records = [r for r in caplog.records if getattr(r, "open_result", None)]
    assert records[-1].open_result == "timeout"


def test_save_failure_terminates_powerpoint(tmp_path, caplog):
    caplog.set_level(logging.INFO)
    working = _working_pptx(tmp_path)
    presentation = FakePresentation(
        working, saved=False, save_error=RuntimeError("disk full")
    )
    app = FakeApp(FakePresentations(presentation))
    killed: list[int] = []
    with pytest.raises(PowerPointAutomationError) as exc:
        _service(tmp_path, app, killed).open(working)
    assert (
        exc.value.user_message
        == "The repaired PowerPoint file could not be saved."
    )
    assert killed == [1]
    records = [r for r in caplog.records if getattr(r, "open_result", None)]
    assert records[-1].open_result == "save_failed"
    assert records[-1].saved is False
