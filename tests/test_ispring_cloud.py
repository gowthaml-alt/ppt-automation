from pathlib import Path

import pytest

from browser_test.ispring_cloud import (
    ArtifactRun,
    extract_material_id,
    is_destructive_label,
    is_official_ispring_url,
    login_credentials,
    looks_like_login,
    normalize_text,
    pick_official_search_result,
    text_matches,
)


def test_normalize_text_collapses_whitespace_and_case():
    assert normalize_text("  Demo   ACADEMY\n") == "demo academy"


def test_text_matches_ignores_capitalization_and_extra_spaces():
    assert text_matches("Demo  Academy (HTML5)", "demo academy")
    assert text_matches("demo academy", "  DEMO   academy ")
    assert text_matches("Demoacademy", "demo academy")
    assert not text_matches("Biology 101", "demo academy")


def test_project_word_sets_match_reordered_names():
    from browser_test.ispring_cloud import _label_word_set

    assert _label_word_set("PPT migration New") == _label_word_set("new ppt migration")
    assert _label_word_set("PPT Migration") < _label_word_set("new ppt migration")


def test_extract_material_id_from_query_string():
    assert (
        extract_material_id("https://pro.ispringcloud.com/app?contentId=abc-123")
        == "abc-123"
    )
    assert (
        extract_material_id("https://pro.ispringcloud.com/?material_id=5001") == "5001"
    )


def test_extract_material_id_from_path_and_hash_route():
    assert (
        extract_material_id("https://pro.ispringcloud.com/contents/xyz789") == "xyz789"
    )
    assert extract_material_id("https://ispringcloud.com/s/AbC12") == "AbC12"
    assert (
        extract_material_id("https://pro.ispringcloud.com/#/contents/hash-id-9")
        == "hash-id-9"
    )
    assert (
        extract_material_id(
            "https://harshit.ispring.com/app/s?s=project%2F"
            "65dcde53-2c22-11ec-b427-a2fdffd5f86a%2F"
            "45ed05bc-2c20-11ec-b54a-76ca0a89c886"
        )
        == "45ed05bc-2c20-11ec-b54a-76ca0a89c886"
    )


def test_extract_material_id_returns_none_when_absent():
    assert extract_material_id("https://pro.ispringcloud.com/") is None
    assert extract_material_id("https://pro.ispringcloud.com/projects") is None
    assert extract_material_id("not a url") is None


def test_looks_like_login_when_welcome_back_and_password_are_visible():
    assert looks_like_login(
        "https://pro.ispringcloud.com/",
        "Welcome Back!\nRemember me\nForgot Password?\nPassword",
    )
    assert looks_like_login("https://pro.ispringcloud.com/login", "Sign in")
    assert looks_like_login(
        "https://id.ispring.com/login?service=isa",
        "Email\nPassword\nLog in",
    )


def test_looks_like_login_is_false_for_the_content_library():
    assert not looks_like_login(
        "https://pro.ispringcloud.com/#/projects",
        "Projects\nRecent\nStarred\nSearch\nShare",
    )


def test_official_ispring_hosts_are_accepted_and_third_parties_are_not():
    assert is_official_ispring_url("https://pro.ispringcloud.com/")
    assert is_official_ispring_url("https://id.ispring.com/login?service=isa")
    assert is_official_ispring_url("https://harshit.ispring.com/app/s")
    assert is_official_ispring_url(
        "https://www.ispringsolutions.com/ispring-cloud?ref=home"
    )
    assert not is_official_ispring_url("https://en.wikipedia.org/wiki/ISpring")
    assert not is_official_ispring_url("https://duckduckgo.com/?q=ispring")


def test_pick_official_search_result_prefers_the_cloud_app():
    chosen = pick_official_search_result(
        [
            ("iSpring Cloud - Wikipedia", "https://en.wikipedia.org/wiki/ISpring"),
            (
                "Easy-to-use online course builder | iSpring Cloud AI",
                "https://www.ispringsolutions.com/ispring-cloud",
            ),
            ("Login", "https://pro.ispringcloud.com/"),
        ]
    )
    assert chosen == "https://pro.ispringcloud.com/"


def test_pick_official_search_result_returns_none_without_an_official_hit():
    assert (
        pick_official_search_result(
            [("Random blog", "https://example.com/ispring-cloud")]
        )
        is None
    )


@pytest.mark.parametrize(
    "label",
    ["Delete", "Upload", "Publish", "Edit", "Remove", "Move to trash"],
)
def test_destructive_labels_are_blocked(label):
    assert is_destructive_label(label)


@pytest.mark.parametrize("label", ["Share", "demo academy", "Search", "Open"])
def test_read_only_labels_are_allowed(label):
    assert not is_destructive_label(label)


def test_login_credentials_require_both_email_and_password():
    assert login_credentials({}) is None
    assert login_credentials({"ISPRING_CLOUD_EMAIL": "user@example.com"}) is None
    assert login_credentials({"ISPRING_CLOUD_PASSWORD": "secret"}) is None
    email, password = login_credentials(
        {
            "ISPRING_CLOUD_EMAIL": " user@example.com ",
            "ISPRING_CLOUD_PASSWORD": "secret",
        }
    )
    assert email == "user@example.com"
    assert password == "secret"


def test_artifact_run_creates_a_timestamped_directory(tmp_path):
    run = ArtifactRun.create(tmp_path)
    assert run.root.is_dir()
    assert run.root.parent == tmp_path / "ispring-cloud"
    assert run.screenshot_name("Opening the browser") == "opening-the-browser.png"
    log_path = run.write_log("opened chrome")
    assert log_path.exists()
    assert "opened chrome" in log_path.read_text(encoding="utf-8")
