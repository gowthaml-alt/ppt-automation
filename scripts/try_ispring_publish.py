"""Drive one iSpring Cloud publish through the PowerPoint interface.

This is the working draft of the `uia` adapter, runnable on its own so it can
be tested without the pipeline. It expects the deck to be ALREADY OPEN in
PowerPoint and out of Protected View.

    python scripts\\try_ispring_publish.py --institution "Demoacademy" --content-name "testing"

Safety switches:
    --stop-before-publish   do everything except the final Publish click
    --inspect               click nothing; just report what can be found

Every step prints what it found, so a failure says which control was missing
rather than dying silently.
"""

from __future__ import annotations

import argparse
import re
import sys
import time
from dataclasses import dataclass, field

# Control ids seen on iSpring Suite 11.17 (probe, 2026-09-16).
ID_CONTENT_NAME = "7545"
ID_PROJECT = "7570"
ID_BROWSE = "7547"
ID_OK = "1"
ID_CANCEL = "2"
ID_DESTINATION_HEADER = "7632"
ID_PROGRESS_STATUS = "141"
ID_DONE_TEXT = "7807"
ID_DONE_DETAIL = "7808"
ID_DONE_MANAGE = "7801"

PUBLISH_DIALOG = "Publish Presentation"
PROJECT_DIALOG = "Select Project"
PROGRESS_PREFIX = "Generating content"
DONE_WINDOW = "iSpring Suite"
RIBBON_TAB = "iSpring Suite 11"

# Dialogs that can appear uninvited and must be dismissed before we can work.
NUISANCE_TITLES = ("checking for updates", "update", "what's new")


def normalise(text: str) -> str:
    """Ribbon labels carry non-breaking spaces and a BOM. Strip all of it."""
    return re.sub(r"\s+", " ", (text or "").replace("﻿", " ").replace("\xa0", " ")).strip()


class PublishError(RuntimeError):
    pass


@dataclass
class Step:
    name: str
    detail: str = ""
    ok: bool = True


@dataclass
class Report:
    steps: list[Step] = field(default_factory=list)

    def add(self, name: str, detail: str = "", ok: bool = True) -> None:
        self.steps.append(Step(name, detail, ok))
        mark = "ok  " if ok else "FAIL"
        print(f"[{mark}] {name}" + (f": {detail}" if detail else ""), flush=True)


def activate(control, report: Report, what: str) -> None:
    """Press a control the gentlest way that works.

    Invoke and select go through the accessibility layer, which keeps working
    when the desktop is not fully drawn. A real click is the last resort.
    """
    last = "no usable method on this control"
    for method in ("invoke", "select", "click_input"):
        action = getattr(control, method, None)
        if action is None:
            continue
        try:
            action()
            report.add(f"pressed {what}", f"via {method}")
            return
        except Exception as exc:  # noqa: BLE001
            last = f"{method}: {type(exc).__name__}: {exc}"
    raise PublishError(f"could not press {what} ({last})")


def set_text(control, value: str, report: Report, what: str) -> None:
    """Fill a text box. Web boxes often ignore a value set without keystrokes."""
    try:
        control.set_edit_text(value)
        if normalise(control.get_value() if hasattr(control, "get_value") else value) == normalise(value):
            report.add(f"filled {what}", repr(value))
            return
    except Exception:  # noqa: BLE001
        pass
    try:
        control.click_input()
        control.type_keys("^a{DELETE}", set_foreground=False)
        control.type_keys(value, with_spaces=True, set_foreground=False)
        report.add(f"typed {what}", repr(value))
    except Exception as exc:  # noqa: BLE001
        raise PublishError(f"could not fill {what}: {exc}") from exc


def by_auto_id(container, auto_id: str) -> list:
    """Find descendants by automation id.

    pywinauto's descendants() does not accept auto_id, so walk and compare.
    """
    found = []
    try:
        children = container.descendants()
    except Exception:  # noqa: BLE001
        return found
    for control in children:
        try:
            if control.element_info.automation_id == auto_id:
                found.append(control)
        except Exception:  # noqa: BLE001
            continue
    return found


def describe(control) -> str:
    try:
        info = control.element_info
        return (
            f"{info.control_type} {info.name!r} "
            f"auto_id={info.automation_id!r} class={info.class_name!r}"
        )
    except Exception:  # noqa: BLE001
        return repr(control)


def wait_for(predicate, timeout_s: float, poll_s: float = 1.0):
    deadline = time.monotonic() + timeout_s
    last = None
    while time.monotonic() < deadline:
        try:
            found = predicate()
        except Exception as exc:  # noqa: BLE001
            last = exc
            found = None
        if found:
            return found
        time.sleep(poll_s)
    if last is not None:
        raise PublishError(f"timed out after {timeout_s}s ({last})")
    raise PublishError(f"timed out after {timeout_s}s")


def open_in_powerpoint(pptx: str, report: Report) -> None:
    """Open the deck through COM so the script can run from a file name alone.

    The file is opened writable and out of Protected View; a deck that stays in
    Protected View has a dead iSpring ribbon.
    """
    import win32com.client  # type: ignore
    from pathlib import Path

    path = Path(pptx).expanduser().resolve()
    if not path.is_file():
        raise PublishError(f"file not found: {path}")

    # Clear the downloaded-from-internet mark, which is what triggers
    # Protected View.
    try:
        import os

        os.remove(f"{path}:Zone.Identifier")
        report.add("cleared internet mark", path.name)
    except FileNotFoundError:
        pass
    except OSError:
        pass

    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = True
    for presentation in app.Presentations:
        try:
            if Path(str(presentation.FullName)).resolve() == path:
                report.add("deck already open", path.name)
                return
        except Exception:  # noqa: BLE001
            continue
    app.Presentations.Open(str(path), False, False, True)
    report.add("opened deck", str(path))
    time.sleep(3)


def find_powerpoint(report: Report):
    from pywinauto import Application  # type: ignore

    app = Application(backend="uia").connect(path="POWERPNT.EXE", timeout=20)
    window = app.window(class_name="PPTFrameClass")
    window.wait("exists", timeout=20)
    report.add("found PowerPoint", window.window_text())
    return app, window


def dismiss_nuisance_dialogs(window, report: Report) -> None:
    """Close update nags and similar pop-ups that block the ribbon."""
    for child in window.children():
        title = normalise(child.window_text()).lower()
        if not title or not any(word in title for word in NUISANCE_TITLES):
            continue
        for candidate in ("Close", "OK", "Cancel", "Later"):
            try:
                button = child.child_window(title=candidate, control_type="Button")
                if button.exists():
                    button.invoke()
                    report.add("closed pop-up", f"{title!r} via {candidate}")
                    break
            except Exception:  # noqa: BLE001
                continue


def find_ribbon_button(window, label: str):
    wanted = normalise(label).lower()
    for control in window.descendants(control_type="Button"):
        try:
            if normalise(control.window_text()).lower() == wanted:
                return control
        except Exception:  # noqa: BLE001
            continue
    return None


def open_publish_dialog(window, report: Report, inspect: bool):
    tab = window.child_window(title=RIBBON_TAB, control_type="TabItem")
    if not tab.exists():
        raise PublishError(f"ribbon tab {RIBBON_TAB!r} not found — is the add-in loaded?")
    report.add("found ribbon tab", RIBBON_TAB)
    if not inspect:
        activate(tab, report, f"tab {RIBBON_TAB!r}")
        time.sleep(1)

    button = find_ribbon_button(window, "Publish")
    if button is None:
        raise PublishError("Publish button not found on the iSpring ribbon")
    report.add("found Publish button", describe(button))
    if inspect:
        return None

    activate(button, report, "ribbon Publish")
    dialog = wait_for(
        lambda: window.child_window(title=PUBLISH_DIALOG, control_type="Window"),
        timeout_s=60,
    )
    dialog.wait("exists visible", timeout=60)
    report.add("publish dialog open", PUBLISH_DIALOG)
    return dialog


def destination_header(dialog) -> str:
    try:
        return normalise(
            dialog.child_window(auto_id=ID_DESTINATION_HEADER).window_text()
        )
    except Exception:  # noqa: BLE001
        return ""


def choose_cloud_destination(dialog, report: Report) -> None:
    """The destination buttons on the left have no names, so go by the header."""
    header = destination_header(dialog)
    if "cloud" in header.lower():
        report.add("destination already correct", header)
        return
    buttons = [
        control
        for control in dialog.descendants(control_type="Button")
        if not normalise(control.window_text())
    ]
    report.add("destination buttons", f"{len(buttons)} unnamed buttons on the left")
    for index, button in enumerate(buttons):
        try:
            button.click_input()
        except Exception:  # noqa: BLE001
            continue
        time.sleep(1.5)
        header = destination_header(dialog)
        if "cloud" in header.lower():
            report.add("selected iSpring Cloud", f"left button #{index + 1}: {header}")
            return
    raise PublishError(
        f"could not switch to iSpring Cloud (header still {header!r})"
    )


def visible_rect(control):
    """Return the control's rectangle, or None when it has no size.

    Rows scrolled out of the picker still exist in the tree, but with a
    zero-size rectangle. That is how we tell "off screen" from "missing".
    """
    try:
        rect = control.rectangle()
    except Exception:  # noqa: BLE001
        return None
    if rect.width() <= 0 or rect.height() <= 0:
        return None
    return rect


def find_named(container, wanted: str, exact: bool = True):
    """Every element in the picker whose name matches, on screen or not."""
    target = normalise(wanted).lower()
    found = []
    for control in container.descendants():
        try:
            name = normalise(control.window_text()).lower()
        except Exception:  # noqa: BLE001
            continue
        if not name:
            continue
        if (name == target) if exact else (target in name):
            found.append(control)
    return found


def scroll_into_view(control, picker, report: Report):
    """Get a row on screen: ask nicely first, then scroll the list."""
    rect = visible_rect(control)
    if rect is not None:
        return rect

    for attempt in ("iface_scroll_item", "set_focus"):
        try:
            if attempt == "iface_scroll_item":
                control.iface_scroll_item.ScrollIntoView()
            else:
                control.set_focus()
            time.sleep(0.6)
            rect = visible_rect(control)
            if rect is not None:
                report.add("scrolled row into view", f"via {attempt}")
                return rect
        except Exception:  # noqa: BLE001
            continue

    # Last resort: turn the wheel over the list and watch for the row to appear.
    surface = None
    for pane in picker.descendants(control_type="Pane"):
        pane_rect = visible_rect(pane)
        if pane_rect is not None and pane_rect.height() > 200:
            surface = pane
            break
    if surface is None:
        surface = picker
    for turn in range(80):
        try:
            surface.wheel_mouse_input(wheel_dist=-3)
        except Exception as exc:  # noqa: BLE001
            raise PublishError(f"could not scroll the project list: {exc}") from exc
        time.sleep(0.25)
        rect = visible_rect(control)
        if rect is not None:
            report.add("scrolled row into view", f"after {turn + 1} wheel turns")
            return rect
    raise PublishError("scrolled to the end of the list and the row never appeared")


def expand_branch(picker, label: str, report: Report) -> bool:
    """Open a folder in the tree, e.g. the parent the institutions live under."""
    for control in find_named(picker, label):
        rect = visible_rect(control)
        if rect is None:
            try:
                rect = scroll_into_view(control, picker, report)
            except PublishError:
                continue
        try:
            control.double_click_input()
            report.add("expanded branch", label)
            time.sleep(2)
            return True
        except Exception:  # noqa: BLE001
            continue
    return False


def pick_project(
    dialog,
    institution: str,
    report: Report,
    parent_folder: str = "",
    list_only: bool = False,
) -> None:
    browse = dialog.child_window(auto_id=ID_BROWSE, control_type="Button")
    if not browse.exists():
        raise PublishError("Browse button (7547) not found on the publish dialog")
    activate(browse, report, "Browse")

    picker = wait_for(
        lambda: dialog.child_window(title=PROJECT_DIALOG, control_type="Window"),
        timeout_s=60,
    )
    picker.wait("exists visible", timeout=60)
    report.add("project picker open", PROJECT_DIALOG)

    if list_only:
        names = []
        for control in picker.descendants():
            try:
                name = normalise(control.window_text())
            except Exception:  # noqa: BLE001
                continue
            if name:
                on_screen = "visible" if visible_rect(control) else "off-screen"
                names.append(f"    {on_screen:>10}  {name}")
        report.add("items in the picker", f"{len(names)}")
        print("\n".join(names[:500]))
        if len(names) > 500:
            print(f"    ... and {len(names) - 500} more")
        return

    matches = find_named(picker, institution)
    if not matches and parent_folder:
        report.add("institution not in tree yet", f"opening {parent_folder!r}")
        expand_branch(picker, parent_folder, report)
        matches = find_named(picker, institution)
    if not matches:
        near = find_named(picker, institution, exact=False)
        hint = ", ".join(normalise(c.window_text()) for c in near[:5]) or "nothing similar"
        raise PublishError(
            f"{institution!r} is not in the project tree (closest: {hint}). "
            "Run again with --list-projects to see the exact names."
        )

    report.add("found project row", describe(matches[0]))
    rect = scroll_into_view(matches[0], picker, report)
    try:
        matches[0].click_input()
        report.add("clicked project row", institution)
    except Exception as exc:  # noqa: BLE001
        raise PublishError(f"could not click the project row at {rect}: {exc}") from exc

    select = picker.child_window(auto_id=ID_OK, control_type="Button")
    if not select.exists():
        raise PublishError("Select button (1) not found in the project picker")
    activate(select, report, "Select")
    wait_for(lambda: not picker.exists(), timeout_s=30)
    report.add("project chosen", institution)


def wait_for_completion(window, report: Report, timeout_s: float) -> None:
    """Wait for the finished window.

    The progress window's status text stops updating and keeps reading
    "Uploading the presentation" after the upload is done, so it must not be
    used as the finish signal. The 'Publishing is complete!' window is.
    """

    def done():
        for control in by_auto_id(window, ID_DONE_TEXT):
            text = normalise(control.window_text())
            if text:
                return control
        return None

    def status() -> str:
        for control in by_auto_id(window, ID_PROGRESS_STATUS):
            text = normalise(control.window_text())
            if text:
                return text
        return ""

    deadline = time.monotonic() + timeout_s
    last_status = ""
    while time.monotonic() < deadline:
        finished = done()
        if finished is not None:
            report.add("publishing complete", normalise(finished.window_text()))
            for control in by_auto_id(window, ID_DONE_DETAIL):
                detail = normalise(control.window_text())
                if detail:
                    report.add("  detail", detail)
            return
        current = status()
        if current and current != last_status:
            last_status = current
            report.add("progress", current)
        time.sleep(2)
    raise PublishError(f"publish did not finish within {timeout_s}s (last status: {last_status!r})")


def click_manage_content(window, report: Report) -> bool:
    """Press Manage Content, which opens the material in the default browser."""
    buttons = by_auto_id(window, ID_DONE_MANAGE)
    if not buttons:
        report.add("no Manage Content button", "nothing to press", ok=False)
        return False
    try:
        activate(buttons[0], report, "Manage Content")
        return True
    except PublishError as exc:
        report.add("could not press Manage Content", str(exc), ok=False)
        return False


def embed_from_running_chrome(
    content_name: str, institution: str, cdp: str, report: Report
) -> str:
    """Ask the already-signed-in Chrome for the embed code.

    Attaching to the running browser is what removes the repeated login: the
    session lives in that Chrome, not in a profile we create.
    """
    try:
        from get_ispring_embed import fetch_embed_over_cdp  # type: ignore
    except ImportError as exc:  # noqa: BLE001
        report.add("embed step unavailable", str(exc), ok=False)
        return ""
    try:
        embed = fetch_embed_over_cdp(
            content_name, institution, cdp=cdp, use_open_tab=True
        )
    except Exception as exc:  # noqa: BLE001
        report.add("embed step failed", f"{type(exc).__name__}: {exc}", ok=False)
        return ""
    if embed:
        report.add("embed code", embed[:120] + ("…" if len(embed) > 120 else ""))
    return embed


def read_link_from_browser(window, report: Report, timeout_s: float = 90) -> str:
    """Click 'Manage Content' and read the URL the browser lands on.

    The finish screen shows no link, so this is the only place a URL appears
    without a separate trip to the iSpring Cloud site. Best effort: it returns
    an empty string rather than failing the publish.
    """
    from pywinauto import Desktop  # type: ignore

    buttons = by_auto_id(window, ID_DONE_MANAGE)
    if not buttons:
        report.add("no Manage Content button", "cannot read a link here", ok=False)
        return ""
    before = time.monotonic()
    try:
        activate(buttons[0], report, "Manage Content")
    except PublishError as exc:
        report.add("could not press Manage Content", str(exc), ok=False)
        return ""

    deadline = before + timeout_s
    while time.monotonic() < deadline:
        for browser in Desktop(backend="uia").windows():
            try:
                title = browser.window_text() or ""
            except Exception:  # noqa: BLE001
                continue
            if not any(word in title for word in ("Chrome", "Edge", "Firefox")):
                continue
            for name in ("Address and search bar", "Address bar", "Address field"):
                try:
                    bar = browser.child_window(title=name, control_type="Edit")
                    if not bar.exists():
                        continue
                    value = normalise(bar.get_value())
                    if value and "." in value:
                        report.add("link from browser", value)
                        return value
                except Exception:  # noqa: BLE001
                    continue
        time.sleep(2)
    report.add("no link found in the browser", f"waited {timeout_s:.0f}s", ok=False)
    return ""


def close_finished_dialogs(window, report: Report) -> None:
    for child in window.descendants(control_type="Window"):
        title = normalise(child.window_text())
        if title != DONE_WINDOW and not title.startswith(PROGRESS_PREFIX):
            continue
        try:
            child.close()
            report.add("closed window", title)
        except Exception:  # noqa: BLE001
            try:
                child.child_window(title="Close", control_type="Button").invoke()
                report.add("closed window", f"{title} (via Close button)")
            except Exception:  # noqa: BLE001
                report.add("could not close window", title, ok=False)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Publish the open presentation to iSpring Cloud."
    )
    parser.add_argument("--pptx", default="", help="Deck to open first (optional)")
    parser.add_argument("--institution", required=True, help="Project name in iSpring Cloud")
    parser.add_argument(
        "--get-link",
        action="store_true",
        help="After publishing, press Manage Content and read the browser's URL",
    )
    parser.add_argument(
        "--embed-after",
        action="store_true",
        help=(
            "After publishing, press Manage Content and then drive that same "
            "signed-in Chrome to Share > Embed and print the iframe"
        ),
    )
    parser.add_argument(
        "--cdp",
        default="http://127.0.0.1:9222",
        help="Debug address of the running Chrome to attach to",
    )
    parser.add_argument("--content-name", default="", help="Content name to set")
    parser.add_argument(
        "--timeout",
        type=float,
        default=1800,
        help="Seconds to wait for the publish to finish (default 1800)",
    )
    parser.add_argument(
        "--parent-folder",
        default="PPT Migration",
        help=(
            "Folder the institutions sit under, opened first if the name is "
            "not already in the tree (default: 'PPT Migration')"
        ),
    )
    parser.add_argument(
        "--list-projects",
        action="store_true",
        help="Open the picker, print every name in it, then stop",
    )
    parser.add_argument("--stop-before-publish", action="store_true")
    parser.add_argument("--inspect", action="store_true", help="Click nothing; only report")
    args = parser.parse_args(argv)

    if sys.platform != "win32":
        print("This needs the Windows machine.", file=sys.stderr)
        return 2
    try:
        import pywinauto  # noqa: F401
    except ImportError:
        print("pip install pywinauto==0.6.8", file=sys.stderr)
        return 2

    report = Report()
    link = ""
    try:
        if args.pptx:
            open_in_powerpoint(args.pptx, report)
        _app, window = find_powerpoint(report)
        dismiss_nuisance_dialogs(window, report)
        dialog = open_publish_dialog(window, report, args.inspect)
        if dialog is None:
            print("\nINSPECT ONLY — nothing was clicked.")
            return 0

        choose_cloud_destination(dialog, report)

        if args.content_name:
            field_ = dialog.child_window(auto_id=ID_CONTENT_NAME, control_type="Edit")
            if not field_.exists():
                raise PublishError("Content name box (7545) not found")
            set_text(field_, args.content_name, report, "content name")

        pick_project(
            dialog,
            args.institution,
            report,
            parent_folder=args.parent_folder,
            list_only=args.list_projects,
        )
        if args.list_projects:
            print("\nLISTED ONLY — nothing was published.")
            return 0

        if args.stop_before_publish:
            print("\nSTOPPED BEFORE PUBLISH — the dialog is set up, nothing was sent.")
            return 0

        publish = dialog.child_window(auto_id=ID_OK, control_type="Button")
        if not publish.exists():
            raise PublishError("Publish button (1) not found on the publish dialog")
        started = time.monotonic()
        activate(publish, report, "Publish")
        wait_for_completion(window, report, args.timeout)
        report.add("elapsed", f"{time.monotonic() - started:.0f}s")
        if args.embed_after:
            if click_manage_content(window, report):
                time.sleep(10)  # let the browser open and settle
                embed = embed_from_running_chrome(
                    args.content_name or args.institution,
                    args.institution,
                    args.cdp,
                    report,
                )
                if embed:
                    match = re.search(r"""\bsrc\s*=\s*["']([^"']+)["']""", embed)
                    link = match.group(1) if match else embed
        elif args.get_link:
            link = read_link_from_browser(window, report)
        close_finished_dialogs(window, report)
    except PublishError as exc:
        print(f"\nPUBLISH FAILED: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # noqa: BLE001
        print(f"\nPUBLISH FAILED (unexpected): {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1

    print("\nPUBLISH OK")
    if link:
        print(f"iframe_url={link}")
    elif args.get_link or args.embed_after:
        print("iframe_url= (not found — see the FAIL lines above)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
