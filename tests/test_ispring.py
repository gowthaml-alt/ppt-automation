from pathlib import Path

import pytest

from config.settings import Settings
from publisher.ispring import (
    CloudPublisher,
    FakePublisher,
    NotConfiguredPublisher,
    get_publisher,
)
from publisher.ispring_cloud import normalise
from utils.exceptions import ISpringNotConfiguredError


def _settings(tmp_path: Path, **overrides) -> Settings:
    """Settings for a test, ignoring any .env on the machine running it.

    Without _env_file=None these tests read the developer's own .env, so a
    machine with ISPRING_ADAPTER=fake would fail the default-value test.
    """
    values = {
        "app_env": "development",
        "backend_base_url": "https://backend.example.com",
        "temp_root": str(tmp_path / "jobs"),
        "log_root": str(tmp_path / "logs"),
    }
    values.update(overrides)
    return Settings(_env_file=None, **values)


def test_default_adapter_is_not_configured(tmp_path):
    publisher = get_publisher(_settings(tmp_path))
    assert isinstance(publisher, NotConfiguredPublisher)
    assert publisher.is_available().installed is False
    with pytest.raises(ISpringNotConfiguredError):
        publisher.publish(tmp_path / "a.pptx", tmp_path / "out", 60)


def test_fake_adapter_copies_the_fixture(tmp_path):
    publisher = get_publisher(_settings(tmp_path, ispring_adapter="fake"))
    assert isinstance(publisher, FakePublisher)
    assert publisher.is_available().installed is True


def test_uia_adapter_is_the_cloud_publisher(tmp_path):
    publisher = get_publisher(_settings(tmp_path, ispring_adapter="uia"))
    assert isinstance(publisher, CloudPublisher)


def test_uia_adapter_refuses_the_folder_style_publish(tmp_path):
    """Cloud publishing returns a URL, so the folder API must not look usable."""
    publisher = get_publisher(_settings(tmp_path, ispring_adapter="uia"))
    with pytest.raises(ISpringNotConfiguredError):
        publisher.publish(tmp_path / "a.pptx", tmp_path / "out", 60)


def test_uia_adapter_is_unavailable_off_windows(tmp_path, monkeypatch):
    monkeypatch.setattr("sys.platform", "darwin")
    publisher = get_publisher(_settings(tmp_path, ispring_adapter="uia"))
    available, detail = (
        publisher.is_available().installed,
        publisher.is_available().detail,
    )
    assert available is False
    assert "Windows" in detail


def test_removed_adapters_are_rejected_by_settings(tmp_path):
    """vba and cli were never implemented and the probes ruled them out."""
    for dead in ("vba", "cli"):
        with pytest.raises(Exception):
            _settings(tmp_path, ispring_adapter=dead)


def test_ribbon_labels_are_normalised():
    """iSpring's ribbon labels carry non-breaking spaces and a BOM."""
    assert normalise("\xa0\xa0Publish\xa0\xa0﻿") == "Publish"
    assert normalise("  iSpring   Suite 11 ") == "iSpring Suite 11"
    assert normalise("") == ""


class _FakeTab:
    def __init__(self, name):
        self._name = name

    def window_text(self):
        return self._name


class _FakeRibbon:
    def __init__(self, names):
        self._names = names

    def descendants(self, control_type=None):
        return [_FakeTab(name) for name in self._names]


def test_the_free_edition_tab_is_accepted():
    from publisher.ispring_cloud import find_ispring_tab

    window = _FakeRibbon(["Home", "Insert", "View", "iSpring Free 11"])
    tab, name = find_ispring_tab(window)
    assert tab is not None
    assert name == "iSpring Free 11"


def test_suite_wins_when_both_tabs_are_present():
    from publisher.ispring_cloud import find_ispring_tab

    window = _FakeRibbon(["Home", "iSpring Free 11", "iSpring Suite 11"])
    _tab, name = find_ispring_tab(window)
    assert name == "iSpring Suite 11"


def test_no_ispring_tab_is_reported_as_missing():
    from publisher.ispring_cloud import find_ispring_tab

    tab, name = find_ispring_tab(_FakeRibbon(["Home", "Insert", "Review"]))
    assert tab is None and name == ""


def test_a_free_only_ribbon_is_not_ready_to_publish():
    from publisher.ispring_cloud import has_addin_tab

    window = _FakeRibbon(["Home", "Insert", "iSpring Free 11"])
    # Free alone means Office disabled Suite: the repair has to run.
    assert has_addin_tab(window, timeout_s=0) is False
    # ...but the publish still goes ahead with it as a last resort.
    assert has_addin_tab(window, timeout_s=0, require_suite=False) is True


def test_a_suite_ribbon_is_ready():
    from publisher.ispring_cloud import has_addin_tab

    window = _FakeRibbon(["Home", "iSpring Suite 11"])
    assert has_addin_tab(window, timeout_s=0) is True


def test_republish_confirms_overwrite_of_the_same_job_id():
    from publisher.ispring_cloud import is_overwrite_prompt, preferred_overwrite_button

    assert is_overwrite_prompt(
        "A presentation with this name already exists. Do you want to replace it?"
    )
    assert is_overwrite_prompt("Content already exists. Overwrite?")
    assert not is_overwrite_prompt("Publishing is complete!")
    assert preferred_overwrite_button(["Cancel", "Yes"]) == "Yes"
    assert preferred_overwrite_button(["Replace", "No"]) == "Replace"


# --- which folder a clashing institution name lands in ----------------------

PARENTS = ["PPT Migration", "PPT migration New"]


def _choose(placed, parents=PARENTS, institution="College Dekho"):
    from publisher.ispring_cloud import choose_match

    return choose_match(placed, parents, institution)


def test_clash_takes_the_first_parent_listed():
    """College Dekho has a folder under both. The order decides, not the tree."""
    chosen = _choose([("PPT migration New", "new"), ("PPT Migration", "old")])
    assert chosen == ["old"]


def test_reversed_order_takes_the_other_parent():
    chosen = _choose(
        [("PPT migration New", "new"), ("PPT Migration", "old")],
        parents=list(reversed(PARENTS)),
    )
    assert chosen == ["new"]


def test_a_flat_tree_keeps_its_matches():
    """No row knows its parent: publish anyway rather than claim it is missing."""
    assert _choose([("", "row")]) == ["row"]


def test_nothing_found_stays_nothing():
    assert _choose([]) == []


# --- the embed has to fill its box, not sit at 1280x720 --------------------


def _fill(embed):
    from publisher.ispring_cloud import fill_container

    return fill_container(embed)


def test_fixed_size_becomes_full_size():
    """What iSpring actually hands back."""
    out = _fill(
        '<iframe src="https://harshit.ispring.com/app/embed-player/abc" '
        'width="1280" height="720" frameborder="0" allowfullscreen></iframe>'
    )
    assert 'width="100%"' in out
    assert 'height="100%"' in out
    assert "1280" not in out and "720" not in out
    # everything else survives
    assert 'src="https://harshit.ispring.com/app/embed-player/abc"' in out
    assert "allowfullscreen" in out
    assert 'frameborder="0"' in out


def test_sizes_are_added_when_ispring_leaves_them_out():
    out = _fill('<iframe src="https://x/y" allowfullscreen></iframe>')
    assert 'width="100%"' in out and 'height="100%"' in out


def test_unquoted_sizes_are_replaced_too():
    out = _fill("<iframe src=https://x/y width=1280 height=720></iframe>")
    assert 'width="100%"' in out and 'height="100%"' in out
    assert "1280" not in out


def test_only_the_iframe_is_touched():
    """A width on something else in the snippet is none of our business."""
    out = _fill(
        '<div width="500"><iframe src="https://x/y" width="1280"></iframe></div>'
    )
    assert '<div width="500">' in out
    assert 'width="100%"' in out
