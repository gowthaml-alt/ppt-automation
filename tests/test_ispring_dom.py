"""The website half, against a copy of iSpring Cloud's own markup.

The HTML below is trimmed from the live library and share popup
(harshit.ispring.com): the same data-at attributes, the same nesting, the
same clipped embed code split across three spans. If iSpring changes its
markup these tests fail, which is the point — better here than on a job.

Skipped when Playwright or its browser is not installed.
"""

from __future__ import annotations

import pytest

playwright_api = pytest.importorskip("playwright.sync_api")

LIBRARY_HTML = """
<!doctype html><html><head><title>iSpring Cloud</title></head><body>
<input data-at="id=search-global-input" type="text" placeholder="Search&hellip;">
<table><tbody data-at="id=table-body">
<tr data-at="selected=false;id=row-ed93084a-b27c-11f1-99fa-beae51271b03">
  <td><a data-at="id=content-item-title">FA 4</a></td>
  <td>Presentation</td><td>PPT Migration</td>
  <td><button data-at="id=table-shortcut-menu-button">&middot;&middot;&middot;</button></td>
</tr>
<tr data-at="selected=false;id=row-bf6ad5a6-b1cd-11f1-b7a8-de7f16e7eda0">
  <td><a data-at="id=content-item-title">20208952-FA_4 [Repaired]</a></td>
  <td>Presentation</td><td>PPT migration New</td>
</tr>
</tbody></table>
</body></html>
"""

SHARE_HTML = """
<!doctype html><html><head><title>iSpring Cloud</title></head><body>
<div data-at="id=uikit-layer-popup"><div data-at="id=sharing-popup">
  <div data-at="id=sharing-content-access-toggle;state=false"><input type="checkbox"></div>
  <div data-at="id=sharing-content-link-field;state=disabled">
    <input data-at="id=input" type="text"
      value="https://harshit.ispring.com/app/preview/ed93084a-b27c-11f1-99fa-beae51271b03">
  </div>
  <div data-at="id=sharing-embed-code-field"><span>&lt;iframe src="</span><span>https://harshit.ispring.com/app/embed-player/ed93084a-b27c-11f1-99fa-beae51271b03</span><span>" width="560" height="315" frameborder="0"&gt;&lt;/iframe&gt;</span></div>
  <div data-at="id=sharing-content-password-access-toggle;state=false"><input type="checkbox"></div>
  <button data-at="id=customize-cover-button">Edit cover image</button>
  <button data-at="id=popup-close-button">X</button>
</div></div></body></html>
"""

PUBLIC_SHARE_HTML = SHARE_HTML.replace(
    "id=sharing-content-access-toggle;state=false",
    "id=sharing-content-access-toggle;state=true",
)


@pytest.fixture(scope="module")
def browser():
    with playwright_api.sync_playwright() as playwright:
        try:
            instance = playwright.chromium.launch()
        except Exception as exc:  # noqa: BLE001
            pytest.skip(f"no browser available: {exc}")
        yield instance
        instance.close()


def _page(browser, html: str, tmp_path, name: str):
    target = tmp_path / name
    target.write_text(html, encoding="utf-8")
    page = browser.new_page()
    page.goto(target.as_uri())
    return page


def test_the_exact_title_wins_over_the_similar_one(browser, tmp_path):
    from publisher.ispring_cloud import find_row

    page = _page(browser, LIBRARY_HTML, tmp_path, "library.html")
    row = find_row(page, "FA 4", scrolls=1)
    assert row is not None
    # Not the "20208952-FA_4 [Repaired]" row left over from an earlier run.
    assert "ed93084a" in (row.get_attribute("data-at") or "")
    page.close()


def test_row_titles_are_read_in_order(browser, tmp_path):
    from publisher.ispring_cloud import row_titles

    page = _page(browser, LIBRARY_HTML, tmp_path, "library2.html")
    assert row_titles(page) == ["FA 4", "20208952-FA_4 [Repaired]"]
    page.close()


def test_the_clipped_embed_code_is_read_whole(browser, tmp_path):
    from publisher.ispring_cloud import read_embed_code

    page = _page(browser, SHARE_HTML, tmp_path, "share.html")
    embed = read_embed_code(page)
    assert embed.startswith("<iframe")
    assert "/app/embed-player/ed93084a-b27c-11f1-99fa-beae51271b03" in embed
    page.close()


def test_link_sharing_state_comes_from_the_toggle(browser, tmp_path):
    from publisher.ispring_cloud import is_public, toggle_state

    off = _page(browser, SHARE_HTML, tmp_path, "share_off.html")
    assert toggle_state(off) is False
    assert is_public(off) is False
    off.close()

    on = _page(browser, PUBLIC_SHARE_HTML, tmp_path, "share_on.html")
    assert toggle_state(on) is True
    assert is_public(on) is True
    on.close()
