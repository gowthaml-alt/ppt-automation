"""Get the embed iframe for a material already published to iSpring Cloud.

Walks the site the way you do by hand: open the library, find the material,
open its three-dot menu, click Share, open the Embed section, and read the
iframe code out of the box.

    python scripts\\get_ispring_embed.py --material "testing" --institution "Demoacademy"

It reads the embed text rather than pressing Copy, because the clipboard
cannot be checked and the text can. Pressing Copy is kept as a fallback.

Screenshots and page dumps land in test-artifacts/ispring-embed/<timestamp>/,
so a failure can be diagnosed from what the page actually showed.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from browser_test.ispring_cloud import (  # noqa: E402
    ArtifactRun,
    ISpringCloudProbeError,
    _dump_page,
    _frames,
    _save_screenshot,
    active_page,
    dismiss_cookie_dialogs,
    launch_persistent_chrome,
    login_credentials,
    maybe_wait_for_login,
    name_pattern,
    open_cloud_library,
    open_ispring_cloud,
    wait_for_app_ready,
    wait_for_visible_text,
)

IFRAME_RE = re.compile(r"<iframe\b[^>]*>(?:.*?</iframe>)?", re.I | re.S)
SRC_RE = re.compile(r"""\bsrc\s*=\s*["']([^"']+)["']""", re.I)

MENU_BUTTON_RE = re.compile(r"more|menu|action|option|\.\.\.|…", re.I)
SHARE_RE = re.compile(r"^\s*share\b", re.I)
EMBED_RE = re.compile(r"embed", re.I)


POPUP_SELECTORS = (
    '[data-at*="uikit-layer-popup"]',
    '[role="dialog"]',
    '[class*="uikit-layer"]',
    '[class*="popup"]',
    '[class*="modal"]',
)


def popup_root(page):
    """The Share popup itself.

    Everything in the dialog must be looked for inside this element. Searching
    the whole page finds matching text in the library list behind the popup,
    and the clicks then land on that list instead of the dialog.
    """
    for frame in _frames(page):
        for selector in POPUP_SELECTORS:
            try:
                locator = frame.locator(selector)
                count = locator.count()
            except Exception:  # noqa: BLE001
                continue
            for index in range(min(count, 5)):
                item = locator.nth(index)
                try:
                    if item.is_visible():
                        return item
                except Exception:  # noqa: BLE001
                    continue
    return None


def scope(page):
    """The popup when one is open, otherwise the whole page."""
    root = popup_root(page)
    return root if root is not None else page


def describe_popup(page, limit: int = 6000) -> str:
    """The popup's markup, for when a control cannot be reached."""
    root = popup_root(page)
    if root is None:
        return "<no popup element found>"
    try:
        return root.evaluate("el => el.outerHTML")[:limit]
    except Exception as exc:  # noqa: BLE001
        return f"<could not read popup html: {exc}>"


def safe_click(locator, what: str = "") -> bool:
    """Click something the page has covered with a styled overlay.

    iSpring's controls are real inputs hidden under decorative divs, so the
    normal click is refused ("subtree intercepts pointer events"). Force skips
    that check; dispatching the event needs no coordinates at all.
    """
    attempts = (
        ("normal", lambda: locator.click(timeout=4000)),
        ("force", lambda: locator.click(force=True, timeout=4000)),
        ("dispatch", lambda: locator.dispatch_event("click")),
        ("check", lambda: locator.check(force=True, timeout=4000)),
    )
    for name, action in attempts:
        try:
            action()
            if name != "normal":
                print(f"[ok  ] clicked {what or 'control'} ({name})")
            return True
        except Exception:  # noqa: BLE001
            continue
    print(f"[FAIL] could not click {what or 'control'} by any method")
    return False


def visible_or_none(locator):
    try:
        count = locator.count()
    except Exception:  # noqa: BLE001
        return None
    for index in range(count):
        item = locator.nth(index)
        try:
            if item.is_visible():
                return item
        except Exception:  # noqa: BLE001
            continue
    return None


def click_by_role(page, pattern, roles=("button", "menuitem", "link", "tab"), what=""):
    for frame in _frames(page):
        for role in roles:
            found = visible_or_none(frame.get_by_role(role, name=pattern))
            if found is not None:
                if safe_click(found, f"{what or pattern.pattern} ({role})"):
                    return True
    found = None
    for frame in _frames(page):
        found = visible_or_none(frame.get_by_text(pattern))
        if found is not None:
            if safe_click(found, f"{what or pattern.pattern} (text)"):
                return True
    return False


def centre_mouse(page) -> None:
    try:
        size = page.viewport_size or {"width": 1200, "height": 800}
        page.mouse.move(size["width"] / 2, size["height"] / 2)
    except Exception:  # noqa: BLE001
        pass


def scroll_hunt(page, name: str, tries: int = 40):
    """Find a row by name, scrolling the list until it shows up.

    The library is a long list, not a search box, so the only way to reach a
    row further down is to keep scrolling and looking.
    """
    pattern = name_pattern(name)
    for attempt in range(tries):
        for frame in _frames(page):
            found = visible_or_none(frame.get_by_text(pattern))
            if found is not None:
                try:
                    found.scroll_into_view_if_needed(timeout=3000)
                except Exception:  # noqa: BLE001
                    pass
                if attempt:
                    print(f"[ok  ] found {name!r} after {attempt} scrolls")
                else:
                    print(f"[ok  ] found {name!r} without scrolling")
                return found
        centre_mouse(page)
        try:
            page.mouse.wheel(0, 700)
        except Exception:  # noqa: BLE001
            break
        page.wait_for_timeout(400)
    return None


def enter_folder(page, folder: str) -> bool:
    """Open the institution's folder, where its materials live."""
    row = scroll_hunt(page, folder)
    if row is None:
        return False
    try:
        row.dblclick(timeout=4000)
        page.wait_for_timeout(2500)
        print(f"[ok  ] opened folder {folder!r} (dblclick)")
        return True
    except Exception:  # noqa: BLE001
        pass
    if safe_click(row, f"folder {folder!r}"):
        page.wait_for_timeout(2500)
        return True
    return False


def menu_items(page) -> list[str]:
    names = []
    for frame in _frames(page):
        for role in ("menuitem", "option", "button", "link"):
            locator = frame.get_by_role(role)
            try:
                count = locator.count()
            except Exception:  # noqa: BLE001
                continue
            for index in range(min(count, 40)):
                item = locator.nth(index)
                try:
                    if not item.is_visible():
                        continue
                    label = (item.inner_text() or "").strip()
                except Exception:  # noqa: BLE001
                    continue
                if label and label not in names:
                    names.append(label)
    return names


def open_row_menu(page, material: str) -> bool:
    """Open the three-dot menu on the material's row.

    The button usually only appears once the row is hovered, and its label
    varies, so hover the row first and then take the nearest menu-ish button.
    """
    row = scroll_hunt(page, material)
    if row is None:
        return False
    try:
        row.hover()
        page.wait_for_timeout(500)
        print("[ok  ] hovered the material row")
    except Exception:  # noqa: BLE001
        pass

    try:
        box = row.bounding_box() or {}
    except Exception:  # noqa: BLE001
        box = {}
    top = box.get("y", 0)
    bottom = top + box.get("height", 0)

    # The three dots sit on the same row, to the right of the name. Keeping to
    # the row's own vertical band avoids grabbing a toolbar button elsewhere.
    candidates = []
    for frame in _frames(page):
        for locator in (
            frame.get_by_role("button", name=MENU_BUTTON_RE),
            frame.get_by_role("button"),
        ):
            try:
                count = locator.count()
            except Exception:  # noqa: BLE001
                continue
            for index in range(min(count, 60)):
                item = locator.nth(index)
                try:
                    if not item.is_visible():
                        continue
                    item_box = item.bounding_box() or {}
                except Exception:  # noqa: BLE001
                    continue
                middle = item_box.get("y", 0) + item_box.get("height", 0) / 2
                if box and not (top - 4 <= middle <= bottom + 4):
                    continue
                candidates.append((item_box.get("x", 0), item))
            if candidates:
                break
        if candidates:
            break
    if not candidates:
        print("[FAIL] no button on the material's row")
        return False

    # Right-most button on the row is the three-dot menu.
    candidates.sort(key=lambda pair: pair[0])
    if not safe_click(candidates[-1][1], "three-dot menu"):
        return False
    page.wait_for_timeout(1200)
    items = menu_items(page)
    print(f"[ok  ] opened the row menu; items: {items}")
    return True


PUBLIC_TOGGLE_RE = re.compile(r"viewable via link|public|share.*link", re.I)
NOT_PUBLIC_RE = re.compile(r"not publicly accessible", re.I)


def is_on(control) -> bool | None:
    """Read a switch's state, whichever way the page reports it."""
    for attribute in ("aria-checked", "aria-pressed", "data-checked"):
        try:
            value = control.get_attribute(attribute)
        except Exception:  # noqa: BLE001
            continue
        if value is not None:
            return value.lower() in ("true", "1", "on", "yes")
    try:
        return bool(control.is_checked())
    except Exception:  # noqa: BLE001
        return None


def is_public(page) -> bool:
    """True only when the dialog no longer says the content is private."""
    return visible_or_none(scope(page).get_by_text(NOT_PUBLIC_RE)) is None


def ancestors(locator, levels: int = 4):
    """The control and its wrappers, outermost last.

    The real input is covered by styled divs, so the thing that actually takes
    the click is usually a parent, not the input itself.
    """
    chain = [locator]
    current = locator
    for _ in range(levels):
        try:
            current = current.locator("xpath=..")
            chain.append(current)
        except Exception:  # noqa: BLE001
            break
    return chain


def toggle_candidates(page):
    """Controls in the popup that might be the 'Make viewable via link' switch.

    Only the first switch in the popup is wanted: the second one is 'Restrict
    with password', which must not be touched.
    """
    root = scope(page)
    found = []
    for getter in (
        lambda: root.get_by_role("switch", name=PUBLIC_TOGGLE_RE),
        lambda: root.get_by_role("checkbox", name=PUBLIC_TOGGLE_RE),
        lambda: root.get_by_role("switch"),
        lambda: root.get_by_role("checkbox"),
        lambda: root.locator("input[type=checkbox]"),
    ):
        try:
            locator = getter()
            count = locator.count()
        except Exception:  # noqa: BLE001
            continue
        if count:
            # First one only: the password switch sits lower in the dialog.
            found.append(locator.first)
    return found


def press_space_on(locator, page, what: str) -> bool:
    """Focus the control and press Space.

    A real checkbox toggles itself on Space and fires its change event, so
    this works even when the input is invisible under a styled switch.
    """
    try:
        locator.focus()
        page.keyboard.press("Space")
        return True
    except Exception:  # noqa: BLE001
        return False


def js_click(locator, what: str) -> bool:
    """Call .click() on the element inside the page.

    Unlike a mouse click this ignores whatever is drawn on top, and unlike
    dispatch_event it also performs the checkbox's own default action.
    """
    try:
        locator.evaluate("el => el.click()")
        return True
    except Exception:  # noqa: BLE001
        return False


def click_switch_beside_off(page) -> bool:
    """Click the visible switch, found by the word 'Off' printed next to it.

    Scoped to the popup: the library behind it also has rows with matching
    text, and clicking those does nothing useful.
    """
    root = scope(page)
    label = visible_or_none(root.get_by_text(PUBLIC_TOGGLE_RE))
    if label is None:
        return False
    try:
        label_box = label.bounding_box()
    except Exception:  # noqa: BLE001
        return False
    if not label_box:
        return False
    row_middle = label_box["y"] + label_box["height"] / 2

    best = None
    locator = root.get_by_text(re.compile(r"^\s*off\s*$", re.I))
    try:
        count = locator.count()
    except Exception:  # noqa: BLE001
        count = 0
    for index in range(min(count, 6)):
        item = locator.nth(index)
        try:
            if not item.is_visible():
                continue
            box = item.bounding_box() or {}
        except Exception:  # noqa: BLE001
            continue
        middle = box.get("y", 0) + box.get("height", 0) / 2
        gap = abs(middle - row_middle)
        if gap < 40 and (best is None or gap < best[0]):
            best = (gap, box)
    if best is None:
        return False

    box = best[1]
    y = box["y"] + box["height"] / 2
    for offset in (30, 45, 60, 20, 75):
        x = box["x"] + box["width"] + offset
        try:
            page.mouse.click(x, y)
            page.wait_for_timeout(1500)
        except Exception:  # noqa: BLE001
            continue
        if is_public(page):
            print(f"[ok  ] clicked the switch {offset}px right of 'Off'")
            return True
    return False


def make_viewable_via_link(page, artifacts) -> bool:
    """Turn on 'Make viewable via link' and prove it went on.

    The embed code is shown even while the content is private, greyed out and
    pointing at something nobody can open. So this must succeed before the
    iframe is worth reading, and it is verified rather than assumed.
    """
    if is_public(page):
        print("[ok  ] content is already viewable via link")
        return True

    for round_number in range(1, 4):
        # Keyboard and in-page click first: neither cares what is drawn on top.
        for candidate in toggle_candidates(page):
            for how, action in (
                ("space", lambda c=candidate: press_space_on(c, page, "toggle")),
                ("js click", lambda c=candidate: js_click(c, "toggle")),
            ):
                if not action():
                    continue
                page.wait_for_timeout(2000)
                if is_public(page):
                    print(f"[ok  ] turned on 'Make viewable via link' ({how})")
                    return True

        if click_switch_beside_off(page) and is_public(page):
            return True

        for candidate in toggle_candidates(page):
            for depth, target in enumerate(ancestors(candidate)):
                try:
                    if not target.is_visible():
                        continue
                except Exception:  # noqa: BLE001
                    continue
                if not safe_click(target, f"link toggle (wrapper {depth})"):
                    continue
                page.wait_for_timeout(2000)
                if is_public(page):
                    print("[ok  ] turned on 'Make viewable via link' (wrapper)")
                    return True
        # The word next to the switch is clickable in some builds.
        for frame in _frames(page):
            label = visible_or_none(frame.get_by_text(re.compile(r"^\s*off\s*$", re.I)))
            if label is not None and safe_click(label, "the Off label"):
                page.wait_for_timeout(2000)
                if is_public(page):
                    print("[ok  ] turned on 'Make viewable via link' (label)")
                    return True
        print(f"[FAIL] toggle still off after attempt {round_number}")
        page.wait_for_timeout(1500)

    _dump_page(page, artifacts, "04c-toggle-stuck-off")
    _save_screenshot(page, artifacts, "04c-toggle-stuck-off")
    print("---- share popup markup ----")
    print(describe_popup(page))
    print("---- end popup markup ----")
    return False


def open_share(page, artifacts) -> bool:
    """Click Share, falling back to the row's right-click menu."""
    if click_by_role(page, SHARE_RE, what="Share"):
        return True
    print(f"[FAIL] no Share in this menu; saw: {menu_items(page)}")
    _dump_page(page, artifacts, "04-no-share")
    return False


def read_embed_code(page) -> str:
    """Pull the iframe snippet out of the share dialog.

    Looks in form fields first, then anywhere on the page, so it does not
    depend on which element iSpring puts the code in.
    """
    root = scope(page)
    for selector in ("textarea", "input[type=text]", "input:not([type])"):
        try:
            locator = root.locator(selector)
            count = locator.count()
        except Exception:  # noqa: BLE001
            count = 0
        for index in range(min(count, 12)):
            try:
                value = locator.nth(index).input_value(timeout=2000)
            except Exception:  # noqa: BLE001
                continue
            if value and "<iframe" in value.lower():
                return value.strip()
    for frame in _frames(page):
        for selector in ("textarea", "input[type=text]", "input:not([type])"):
            try:
                handles = frame.query_selector_all(selector)
            except Exception:  # noqa: BLE001
                continue
            for handle in handles:
                try:
                    value = handle.input_value()
                except Exception:  # noqa: BLE001
                    continue
                if value and "<iframe" in value.lower():
                    return value.strip()
    for frame in _frames(page):
        try:
            content = frame.content()
        except Exception:  # noqa: BLE001
            continue
        # The embed code is displayed as escaped text, so look at the text too.
        try:
            text = frame.inner_text("body")
        except Exception:  # noqa: BLE001
            text = ""
        for blob in (text, content):
            match = IFRAME_RE.search(blob or "")
            if match and "ispring" in match.group(0).lower():
                return match.group(0).strip()
    return ""


def copy_button_fallback(page) -> str:
    """Press a Copy button next to the embed box and read the clipboard."""
    if not click_by_role(page, re.compile(r"^\s*copy\b", re.I), what="Copy"):
        return ""
    try:
        import win32clipboard  # type: ignore

        win32clipboard.OpenClipboard()
        try:
            value = win32clipboard.GetClipboardData()
        finally:
            win32clipboard.CloseClipboard()
        return (value or "").strip()
    except Exception as exc:  # noqa: BLE001
        print(f"[FAIL] could not read the clipboard: {exc}")
        return ""


def report_embed(embed: str) -> int:
    if not embed:
        print("\nEMBED NOT FOUND")
        return 1
    print("\nEMBED OK")
    print(f"embed_code={embed}")
    src = SRC_RE.search(embed)
    if src:
        print(f"iframe_url={src.group(1)}")
    return 0


def looks_logged_out(page) -> bool:
    try:
        from browser_test.ispring_cloud import _visible_text, looks_like_login

        return looks_like_login(page.url, _visible_text(page))
    except Exception:  # noqa: BLE001
        return False


def pick_ispring_page(context, use_open_tab: bool):
    """Find the iSpring tab in an already-running browser.

    The tab opened by 'Manage Content' is usually the last one, and it is
    already on the right material, so prefer the newest match.
    """
    matches = []
    for page in context.pages:
        try:
            url = page.url or ""
        except Exception:  # noqa: BLE001
            continue
        if "ispring" in url.lower():
            matches.append(page)
    if matches:
        return matches[-1]
    if use_open_tab:
        return None
    return context.new_page()


def share_flow(page, context, material: str, institution: str, artifacts) -> str:
    """The part that is the same whether we launched Chrome or attached to it."""
    wait_for_app_ready(page)
    dismiss_cookie_dialogs(page)
    _save_screenshot(page, artifacts, "01-library")

    # No searching: the material is found by walking into the institution's
    # folder and scrolling the list, the way it is done by hand.
    if not wait_for_visible_text(page, name_pattern(material), 5_000):
        if institution:
            if not enter_folder(page, institution):
                _dump_page(page, artifacts, "02-no-folder")
                raise ISpringCloudProbeError(
                    f"could not open the folder {institution!r} in the library"
                )
            wait_for_app_ready(page, timeout_ms=20_000)
    _save_screenshot(page, artifacts, "02-folder")
    _dump_page(page, artifacts, "02-folder")

    if not open_row_menu(page, material):
        _dump_page(page, artifacts, "03-no-row-menu")
        raise ISpringCloudProbeError(
            f"could not open the three-dot menu for {material!r} — is that the "
            "name it was published under?"
        )
    _save_screenshot(page, artifacts, "03-row-menu")

    if not open_share(page, artifacts):
        raise ISpringCloudProbeError("no Share item in the menu")
    page.wait_for_timeout(2000)
    _save_screenshot(page, artifacts, "04-share-dialog")
    _dump_page(page, artifacts, "04-share-dialog")

    if not make_viewable_via_link(page, artifacts):
        raise ISpringCloudProbeError(
            "'Make viewable via link' could not be switched on. The embed code "
            "on screen points at content nobody can open, so it is not worth "
            "returning — see 04c-toggle-stuck-off.png"
        )
    page.wait_for_timeout(2000)
    _save_screenshot(page, artifacts, "04d-after-toggle")
    _dump_page(page, artifacts, "04d-after-toggle")

    # With sharing on, the embed code is usually already on the dialog.
    embed = read_embed_code(page)
    if embed:
        print("[ok  ] embed code was already on the dialog")
        return embed

    click_by_role(page, EMBED_RE, roles=("tab", "button", "link"), what="Embed code")
    page.wait_for_timeout(1500)
    _save_screenshot(page, artifacts, "05-embed")
    _dump_page(page, artifacts, "05-embed")

    embed = read_embed_code(page)
    if not embed:
        print("[FAIL] no iframe in the page text; trying the Copy button")
        embed = copy_button_fallback(page)
    return embed


def fetch_embed_over_cdp(
    material: str,
    institution: str = "",
    *,
    cdp: str = "http://127.0.0.1:9222",
    use_open_tab: bool = True,
    artifacts_dir: str | Path = "",
) -> str:
    """Drive the Chrome that is already running and already signed in.

    Nothing is launched and no profile is created, so the login that was done
    once by hand in that Chrome keeps working for every run after it.
    """
    from playwright.sync_api import sync_playwright

    artifacts = ArtifactRun.create(
        Path(artifacts_dir or (ROOT / "test-artifacts" / "ispring-embed"))
    )
    with sync_playwright() as playwright:
        try:
            browser = playwright.chromium.connect_over_cdp(cdp)
        except Exception as exc:  # noqa: BLE001
            raise ISpringCloudProbeError(
                f"could not attach to Chrome at {cdp}: {exc}. Start Chrome with "
                '--remote-debugging-port=9222 (see scripts\\start_ispring_chrome.cmd).'
            ) from exc
        contexts = browser.contexts
        if not contexts:
            raise ISpringCloudProbeError("attached to Chrome but it has no windows open")
        context = contexts[0]
        page = pick_ispring_page(context, use_open_tab)
        if page is None:
            raise ISpringCloudProbeError(
                "no iSpring tab is open in that Chrome. Publish first so "
                "'Manage Content' opens one, or drop --use-open-tab."
            )
        page.bring_to_front()
        if looks_logged_out(page):
            raise ISpringCloudProbeError(
                "that Chrome is not signed in to iSpring Cloud. Sign in once by "
                "hand in this browser; the session is then reused every run."
            )
        embed = share_flow(page, context, material, institution, artifacts)
        print(f"artifacts={artifacts.root}")
        return embed


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Read the embed iframe for a published material."
    )
    parser.add_argument("--material", required=True, help="Content name as published")
    parser.add_argument(
        "--institution",
        default="",
        help="Folder to search for if the material is not in the first list",
    )
    parser.add_argument("--url", default=os.environ.get("ISPRING_CLOUD_URL", ""))
    parser.add_argument(
        "--artifacts-dir", default=str(ROOT / "test-artifacts" / "ispring-embed")
    )
    parser.add_argument(
        "--profile-dir", default=str(ROOT / "test-artifacts" / "chrome-profile")
    )
    parser.add_argument("--headless", action="store_true")
    parser.add_argument(
        "--cdp",
        default="http://127.0.0.1:9222",
        help=(
            "Attach to a running Chrome at this debug address instead of "
            "launching one. This is what avoids logging in every time. "
            "Pass an empty string to launch a separate Chrome instead."
        ),
    )
    parser.add_argument(
        "--use-open-tab",
        action="store_true",
        help="Use the iSpring tab already open (the one Manage Content opened)",
    )
    args = parser.parse_args(argv)

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("playwright is not installed: pip install -r requirements.txt", file=sys.stderr)
        return 2

    if args.cdp:
        try:
            embed = fetch_embed_over_cdp(
                args.material,
                args.institution,
                cdp=args.cdp,
                use_open_tab=args.use_open_tab,
                artifacts_dir=args.artifacts_dir,
            )
        except ISpringCloudProbeError as exc:
            print(f"\nEMBED FAILED: {exc}", file=sys.stderr)
            return 1
        except Exception as exc:  # noqa: BLE001
            print(
                f"\nEMBED FAILED (unexpected): {type(exc).__name__}: {exc}",
                file=sys.stderr,
            )
            return 1
        return report_embed(embed)

    # Fallback: launch our own Chrome with a saved profile (needs one login).
    artifacts = ArtifactRun.create(Path(args.artifacts_dir))
    supplied = login_credentials(os.environ)
    email = supplied[0] if supplied else None
    password = supplied[1] if supplied else None

    embed = ""
    with sync_playwright() as playwright:
        context = launch_persistent_chrome(
            playwright, Path(args.profile_dir), headless=args.headless
        )
        try:
            page = context.pages[0] if context.pages else context.new_page()
            open_ispring_cloud(page, args.url or None)
            page = active_page(context, page)
            maybe_wait_for_login(
                page, None if supplied else input, email=email, password=password
            )
            page = open_cloud_library(context, page)
            embed = share_flow(page, context, args.material, args.institution, artifacts)
        except ISpringCloudProbeError as exc:
            print(f"\nEMBED FAILED: {exc}", file=sys.stderr)
            print(f"artifacts={artifacts.root}", file=sys.stderr)
            return 1
        except Exception as exc:  # noqa: BLE001
            print(
                f"\nEMBED FAILED (unexpected): {type(exc).__name__}: {exc}",
                file=sys.stderr,
            )
            print(f"artifacts={artifacts.root}", file=sys.stderr)
            return 1
        finally:
            try:
                context.close()
            except Exception:  # noqa: BLE001
                pass

    print(f"artifacts={artifacts.root}")
    return report_embed(embed)


if __name__ == "__main__":
    raise SystemExit(main())
