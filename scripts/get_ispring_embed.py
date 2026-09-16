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
    search_library,
    wait_for_app_ready,
    wait_for_visible_text,
)

IFRAME_RE = re.compile(r"<iframe\b[^>]*>(?:.*?</iframe>)?", re.I | re.S)
SRC_RE = re.compile(r"""\bsrc\s*=\s*["']([^"']+)["']""", re.I)

MENU_BUTTON_RE = re.compile(r"more|menu|action|option|\.\.\.|…", re.I)
SHARE_RE = re.compile(r"^\s*share\b", re.I)
EMBED_RE = re.compile(r"embed", re.I)


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
                found.click()
                print(f"[ok  ] clicked {what or pattern.pattern} ({role})")
                return True
    found = None
    for frame in _frames(page):
        found = visible_or_none(frame.get_by_text(pattern))
        if found is not None:
            found.click()
            print(f"[ok  ] clicked {what or pattern.pattern} (text)")
            return True
    return False


def open_row_menu(page, material: str) -> bool:
    """Open the three-dot menu on the material's row.

    The button usually only appears once the row is hovered, and its label
    varies, so hover the row first and then take the nearest menu-ish button.
    """
    row = None
    for frame in _frames(page):
        candidate = visible_or_none(frame.get_by_text(name_pattern(material)))
        if candidate is not None:
            row = candidate
            break
    if row is None:
        return False
    try:
        row.hover()
        print("[ok  ] hovered the material row")
    except Exception:  # noqa: BLE001
        pass

    try:
        box = row.bounding_box() or {}
        row_middle = box.get("y", 0) + box.get("height", 0) / 2
    except Exception:  # noqa: BLE001
        row_middle = None

    best = None
    best_gap = None
    for frame in _frames(page):
        locator = frame.get_by_role("button", name=MENU_BUTTON_RE)
        try:
            count = locator.count()
        except Exception:  # noqa: BLE001
            continue
        for index in range(count):
            item = locator.nth(index)
            try:
                if not item.is_visible():
                    continue
                if row_middle is None:
                    best = item
                    break
                item_box = item.bounding_box() or {}
                middle = item_box.get("y", 0) + item_box.get("height", 0) / 2
                gap = abs(middle - row_middle)
                if best_gap is None or gap < best_gap:
                    best, best_gap = item, gap
            except Exception:  # noqa: BLE001
                continue
    if best is None:
        return False
    best.click()
    print(f"[ok  ] opened the row menu (gap {best_gap})")
    return True


def read_embed_code(page) -> str:
    """Pull the iframe snippet out of the share dialog.

    Looks in form fields first, then anywhere on the page, so it does not
    depend on which element iSpring puts the code in.
    """
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

    if not wait_for_visible_text(page, name_pattern(material), 6_000):
        search_library(page, material)
        page = active_page(context, page)
        wait_for_app_ready(page, timeout_ms=20_000)
    if not wait_for_visible_text(page, name_pattern(material), 15_000):
        if institution:
            search_library(page, institution)
            wait_for_app_ready(page, timeout_ms=20_000)
            search_library(page, material)
            wait_for_app_ready(page, timeout_ms=20_000)
    _save_screenshot(page, artifacts, "02-material-found")
    _dump_page(page, artifacts, "02-material-found")

    if not open_row_menu(page, material):
        _dump_page(page, artifacts, "03-no-row-menu")
        raise ISpringCloudProbeError(
            f"could not find the three-dot menu for {material!r}"
        )
    page.wait_for_timeout(1200)
    _save_screenshot(page, artifacts, "03-row-menu")

    if not click_by_role(page, SHARE_RE, what="Share"):
        _dump_page(page, artifacts, "04-no-share")
        raise ISpringCloudProbeError("no Share item in the menu")
    page.wait_for_timeout(2000)
    _save_screenshot(page, artifacts, "04-share-dialog")
    _dump_page(page, artifacts, "04-share-dialog")

    click_by_role(page, EMBED_RE, roles=("tab", "button", "link"), what="Embed")
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
