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
    # Ignore any .env on the machine running the tests.
    return Settings(_env_file=None, **values)


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


def test_workspace_is_deleted(tmp_path):
    from worker.cloud_job import remove_workspace

    workspace = tmp_path / "cloud-123"
    (workspace / "nested").mkdir(parents=True)
    (workspace / "source.pptx").write_bytes(b"PK\x03\x04")
    (workspace / "source.repaired.pptx").write_bytes(b"PK\x03\x04")
    assert remove_workspace(workspace) is True
    assert not workspace.exists()


def test_removing_a_missing_workspace_is_fine(tmp_path):
    from worker.cloud_job import remove_workspace

    assert remove_workspace(tmp_path / "never-existed") is True


def test_addin_check_is_a_no_op_off_windows():
    """The check must not explode where there is no registry to read."""
    from powerpoint.addin import check_addin, ensure_addin_enabled

    state = check_addin()
    assert state.checked is False
    assert state.healthy is True
    assert ensure_addin_enabled() is True


def test_addin_state_is_unhealthy_when_office_demoted_it():
    from powerpoint.addin import AddinState

    fine = AddinState(registered=["HKCU\\...\\iSpringSuite11.Connect"])
    assert fine.healthy is True

    demoted = AddinState(
        registered=["HKCU\\...\\iSpringSuite11.Connect"],
        demoted=["HKCU\\...\\iSpringSuite11.Connect (LoadBehavior=2)"],
    )
    assert demoted.healthy is False

    switched_off = AddinState(
        registered=["HKCU\\...\\iSpringSuite11.Connect"],
        disabled_items=[("...\\DisabledItems", "item1")],
    )
    assert switched_off.healthy is False

    missing = AddinState()
    assert missing.healthy is False


class _FakeWindow:
    pass


def _patch_addin_check(monkeypatch, *, tab_states, connected):
    """tab_states is consumed one per has_addin_tab call."""
    import powerpoint.addin as addin_mod
    import publisher.ispring_cloud as cloud_mod

    calls = {"connect": 0}
    states = list(tab_states)

    monkeypatch.setattr(cloud_mod, "find_powerpoint_window", lambda: _FakeWindow())
    monkeypatch.setattr(
        cloud_mod, "has_addin_tab", lambda window, timeout_s=20: states.pop(0)
    )

    def _connect(app):
        calls["connect"] += 1
        return connected

    monkeypatch.setattr(addin_mod, "connect_addin", _connect)
    return calls


def test_addin_ready_when_the_tab_is_there(monkeypatch):
    from worker.cloud_job import addin_ready

    calls = _patch_addin_check(monkeypatch, tab_states=[True], connected=False)
    assert addin_ready(object()) is True
    assert calls["connect"] == 0  # nothing to fix, so nothing is touched


def test_missing_tab_is_switched_on_without_restarting(monkeypatch):
    from worker.cloud_job import addin_ready

    calls = _patch_addin_check(monkeypatch, tab_states=[False, True], connected=True)
    assert addin_ready(object()) is True
    assert calls["connect"] == 1


def test_missing_tab_that_cannot_be_switched_on_asks_for_a_restart(monkeypatch):
    from worker.cloud_job import addin_ready

    _patch_addin_check(monkeypatch, tab_states=[False, False], connected=True)
    assert addin_ready(object()) is False


def test_the_deck_is_published_under_the_queue_id_not_its_name(monkeypatch, tmp_path):
    """The Suite publish name is the id; the cloud cover keeps the name.

    Two materials can share a title, and the browser half then cannot tell
    which row it just published. An id is unique, so it is what goes in the
    publish dialog.
    """
    import worker.cloud_job as job

    seen = {}

    def _publish(pptx, **kwargs):
        seen.update(kwargs)
        return type("R", (), {"iframe_url": "u", "embed_code": "e", "elapsed_s": 1})()

    monkeypatch.setattr(job, "publish_to_cloud", _publish)
    monkeypatch.setattr(job, "inspect", lambda _p: False)
    monkeypatch.setattr(
        job, "open_with_repair", lambda deck, _s: (None, deck, 0)
    )

    source = tmp_path / "deck.pptx"
    source.write_bytes(minimal_pptx_bytes())

    job.run_cloud_job(
        pptx=str(source),
        material_name="Kickoff and Advanced Prompting",
        material_id=20242897,
        job_id=412,
        institution_name="Demoacademy",
        settings=_settings(tmp_path),
    )

    assert seen["content_name"] == "412"            # the Suite dialog
    assert seen["cover_title"] == "Kickoff and Advanced Prompting"  # the cloud


def test_without_a_job_id_the_material_id_is_used(monkeypatch, tmp_path):
    import worker.cloud_job as job

    seen = {}
    monkeypatch.setattr(
        job, "publish_to_cloud",
        lambda pptx, **kw: (seen.update(kw),
                            type("R", (), {"iframe_url": "u", "embed_code": "e", "elapsed_s": 1})())[1],
    )
    monkeypatch.setattr(job, "inspect", lambda _p: False)
    monkeypatch.setattr(job, "open_with_repair", lambda deck, _s: (None, deck, 0))

    source = tmp_path / "deck.pptx"
    source.write_bytes(minimal_pptx_bytes())

    job.run_cloud_job(
        pptx=str(source),
        material_name="Kickoff",
        material_id=20242897,
        institution_name="Demoacademy",
        settings=_settings(tmp_path),
    )
    assert seen["content_name"] == "20242897"


# --- an institution with no folder yet -------------------------------------


def test_a_missing_folder_is_created_and_the_deck_published(monkeypatch, tmp_path):
    """First publish finds no folder, one is made, the second publish works."""
    import worker.cloud_job as job
    from utils.exceptions import ProjectMissingError

    attempts = []
    created = []

    def _publish(pptx, **kwargs):
        attempts.append(kwargs)
        if len(attempts) == 1:
            raise ProjectMissingError("no folder", institution="Brand New College")
        return type("R", (), {"iframe_url": "u", "embed_code": "e", "elapsed_s": 1})()

    monkeypatch.setattr(job, "publish_to_cloud", _publish)
    monkeypatch.setattr(
        job, "ensure_project_folder",
        lambda institution, parent, cdp, **kw: created.append((institution, parent)),
    )
    monkeypatch.setattr(job, "inspect", lambda _p: False)
    monkeypatch.setattr(job, "open_with_repair", lambda deck, _s: (None, deck, 0))

    source = tmp_path / "deck.pptx"
    source.write_bytes(minimal_pptx_bytes())

    result = job.run_cloud_job(
        pptx=str(source),
        material_name="Induction",
        material_id=1,
        job_id=77,
        institution_name="Brand New College",
        settings=_settings(tmp_path),
    )

    assert len(attempts) == 2
    assert created == [("Brand New College", "PPT migration New")]
    assert result.iframe_url == "u"


def test_the_folder_is_created_once_not_in_a_loop(monkeypatch, tmp_path):
    """Still missing after creating it: give up rather than go round again."""
    import worker.cloud_job as job
    from utils.exceptions import ProjectMissingError

    attempts = []

    def _publish(pptx, **kwargs):
        attempts.append(kwargs)
        raise ProjectMissingError("no folder", institution="Brand New College")

    monkeypatch.setattr(job, "publish_to_cloud", _publish)
    monkeypatch.setattr(job, "ensure_project_folder", lambda *a, **k: True)
    monkeypatch.setattr(job, "inspect", lambda _p: False)
    monkeypatch.setattr(job, "open_with_repair", lambda deck, _s: (None, deck, 0))

    source = tmp_path / "deck.pptx"
    source.write_bytes(minimal_pptx_bytes())

    with pytest.raises(ProjectMissingError):
        job.run_cloud_job(
            pptx=str(source),
            material_name="Induction",
            material_id=1,
            job_id=78,
            institution_name="Brand New College",
            settings=_settings(tmp_path),
        )
    assert len(attempts) == 2
