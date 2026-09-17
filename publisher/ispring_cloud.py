"""Publish a deck to iSpring Cloud and come back with its iframe URL.

There is no iSpring API to call: the add-in exposes no automation object, no
macro entry point and no command line (see docs and scripts/probe_ispring.py).
So the desktop half drives the PowerPoint interface, and the website half
drives a Chrome that is already signed in.

Two halves, in order:

1. Desktop  - iSpring Suite ribbon > Publish > iSpring Cloud > pick the
   institution's project > Publish, then wait for "Publishing is complete!".
2. Website  - press Manage Content, attach to the running Chrome, open the
   material's Share popup, switch on "Make viewable via link", and read the
   embed iframe.

Both halves need an interactive desktop: a locked screen has no active
desktop, and interface automation stops working there.
"""

from __future__ import annotations

import logging
import os
import re
import time
from dataclasses import dataclass
from pathlib import Path

from utils.exceptions import ISpringPublishingError

logger = logging.getLogger(__name__)

STAGE = "ispring_publish"

# --- desktop control ids, iSpring Suite 11.17 -------------------------------
ID_CONTENT_NAME = "7545"
ID_PROJECT = "7570"
ID_BROWSE = "7547"
ID_OK = "1"
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

NUISANCE_TITLES = ("checking for updates", "update", "what's new")

# PowerPoint asks before repairing a damaged deck, and the dialog blocks the
# COM call that opened the file — so it is answered from another thread.
REPAIR_PROMPT_RE = re.compile(
    r"repair|recover|unreadable content|found a problem|damaged", re.I
)
REPAIR_BUTTONS = ("Repair", "Yes", "OK", "Open", "Continue")

# --- website selectors ------------------------------------------------------
POPUP_SELECTORS = (
    '[data-at*="uikit-layer-popup"]',
    '[role="dialog"]',
    '[class*="uikit-layer"]',
    '[class*="popup"]',
    '[class*="modal"]',
)
IFRAME_RE = re.compile(r"<iframe\b[^>]*>(?:.*?</iframe>)?", re.I | re.S)
SRC_RE = re.compile(r"""\bsrc\s*=\s*["']([^"']+)["']""", re.I)
MENU_BUTTON_RE = re.compile(r"more|menu|action|option|\.\.\.|…", re.I)
SHARE_RE = re.compile(r"^\s*share\b", re.I)
EMBED_RE = re.compile(r"embed", re.I)
PUBLIC_TOGGLE_RE = re.compile(r"viewable via link|public|share.*link", re.I)
NOT_PUBLIC_RE = re.compile(r"not publicly accessible", re.I)
OFF_RE = re.compile(r"^\s*off\s*$", re.I)

USER_PUBLISH_FAILED = "The presentation could not be published to iSpring Cloud."
USER_LINK_FAILED = "The presentation was published but no share link could be read."


@dataclass(frozen=True)
class CloudPublishResult:
    iframe_url: str
    embed_code: str
    elapsed_s: float


def _note(message: str, **fields) -> None:
    logger.info(message, extra={"stage": STAGE, **fields})


def _warn(message: str, **fields) -> None:
    logger.warning(message, extra={"stage": STAGE, **fields})


def ensure_com() -> None:
    """Initialise COM on this thread.

    Both halves of this module reach COM through different libraries —
    pywin32 for PowerPoint, comtypes underneath pywinauto — and each expects
    COM to be ready on the thread it runs on. Calling this is cheap and safe
    to repeat; skipping it produces "CoInitialize has not been called" at the
    first pywinauto import.
    """
    try:
        import pythoncom  # type: ignore

        pythoncom.CoInitialize()
    except Exception:  # noqa: BLE001
        pass


def normalise(text: str) -> str:
    """Ribbon labels carry non-breaking spaces and a BOM. Strip all of it."""
    cleaned = (text or "").replace("﻿", " ").replace("\xa0", " ")
    return re.sub(r"\s+", " ", cleaned).strip()


# ---------------------------------------------------------------------------
# Desktop half
# ---------------------------------------------------------------------------


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


def activate(control, what: str) -> None:
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
            _note("pressed control", control=what, method=method)
            return
        except Exception as exc:  # noqa: BLE001
            last = f"{method}: {type(exc).__name__}: {exc}"
    raise ISpringPublishingError(
        f"could not press {what} ({last})", user_message=USER_PUBLISH_FAILED
    )


def set_text(control, value: str, what: str) -> None:
    try:
        control.set_edit_text(value)
        _note("filled field", field=what)
        return
    except Exception:  # noqa: BLE001
        pass
    try:
        control.click_input()
        control.type_keys("^a{DELETE}", set_foreground=False)
        control.type_keys(value, with_spaces=True, set_foreground=False)
        _note("typed field", field=what)
    except Exception as exc:  # noqa: BLE001
        raise ISpringPublishingError(
            f"could not fill {what}: {exc}", user_message=USER_PUBLISH_FAILED
        ) from exc


def visible_rect(control):
    """The control's rectangle, or None when it has no size.

    Rows scrolled out of the project tree still exist, but with a zero-size
    rectangle. That is how "off screen" is told apart from "missing".
    """
    try:
        rect = control.rectangle()
    except Exception:  # noqa: BLE001
        return None
    if rect.width() <= 0 or rect.height() <= 0:
        return None
    return rect


def find_named(container, wanted: str, exact: bool = True) -> list:
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


def open_presentation(pptx: Path):
    """Open the deck through COM, out of Protected View.

    A deck in Protected View has a dead iSpring ribbon, and the mark of the
    web is what puts it there.
    """
    import win32com.client  # type: ignore

    ensure_com()
    path = Path(pptx).expanduser().resolve()
    if not path.is_file():
        raise ISpringPublishingError(
            f"file not found: {path}", user_message=USER_PUBLISH_FAILED
        )
    try:
        os.remove(f"{path}:Zone.Identifier")
        _note("cleared mark of the web", file=path.name)
    except (FileNotFoundError, OSError):
        pass

    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = True
    for presentation in app.Presentations:
        try:
            if Path(str(presentation.FullName)).resolve() == path:
                _note("deck already open", file=path.name)
                return app
        except Exception:  # noqa: BLE001
            continue
    app.Presentations.Open(str(path), False, False, True)
    _note("opened deck", file=path.name)
    time.sleep(3)
    return app


def find_powerpoint_window():
    ensure_com()
    from pywinauto import Application  # type: ignore

    app = Application(backend="uia").connect(path="POWERPNT.EXE", timeout=30)
    window = app.window(class_name="PPTFrameClass")
    window.wait("exists", timeout=30)
    _note("found PowerPoint", title=window.window_text())
    return window


def dismiss_nuisance_dialogs(window) -> None:
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
                    _warn("closed an iSpring pop-up", popup=title, button=candidate)
                    break
            except Exception:  # noqa: BLE001
                continue


def answer_repair_prompt() -> bool:
    """Click through PowerPoint's 'repair this file?' dialog.

    Looks across every top-level window rather than under PowerPoint's own,
    because the prompt can appear before the main window is usable.
    """
    try:
        from pywinauto import Desktop  # type: ignore
    except ImportError:
        return False
    answered = False
    try:
        windows = Desktop(backend="uia").windows()
    except Exception:  # noqa: BLE001
        return False
    for window in windows:
        try:
            title = normalise(window.window_text())
        except Exception:  # noqa: BLE001
            continue
        haystack = title
        if not REPAIR_PROMPT_RE.search(haystack):
            try:
                haystack = " ".join(
                    normalise(child.window_text())
                    for child in window.descendants(control_type="Text")[:12]
                )
            except Exception:  # noqa: BLE001
                continue
            if not REPAIR_PROMPT_RE.search(haystack):
                continue
        for label in REPAIR_BUTTONS:
            try:
                button = window.child_window(title=label, control_type="Button")
                if not button.exists():
                    continue
                button.invoke()
                _warn("answered a PowerPoint repair prompt", dialog=title, button=label)
                answered = True
                break
            except Exception:  # noqa: BLE001
                continue
    return answered


class RepairPromptWatcher:
    """Answer repair prompts while a blocking open is in flight.

    PowerPoint's COM Open call does not return while its dialog is up, so the
    dialog has to be clicked from a second thread or the open simply hangs
    until the timeout.
    """

    def __init__(self, poll_s: float = 1.0) -> None:
        self._poll_s = poll_s
        self._stop = None
        self._thread = None
        self.answered = 0

    def __enter__(self) -> "RepairPromptWatcher":
        import threading

        # Import pywinauto here, on the caller's thread, so comtypes sets COM
        # up for the main thread. Letting the watcher thread import it first
        # leaves the main thread without COM once the watcher ends.
        try:
            ensure_com()
            import pywinauto  # noqa: F401
        except Exception:  # noqa: BLE001
            _warn("pywinauto unavailable; repair prompts will not be answered")
            return self

        self._stop = threading.Event()

        def loop() -> None:
            ensure_com()
            try:
                while not self._stop.wait(self._poll_s):
                    try:
                        if answer_repair_prompt():
                            self.answered += 1
                    except Exception:  # noqa: BLE001
                        continue
            finally:
                # Leave this thread's COM as we found it; the main thread's
                # own initialisation must not be affected.
                try:
                    import pythoncom  # type: ignore

                    pythoncom.CoUninitialize()
                except Exception:  # noqa: BLE001
                    pass

        self._thread = threading.Thread(
            target=loop, daemon=True, name="ppt-repair-prompt"
        )
        self._thread.start()
        return self

    def __exit__(self, *exc_info) -> None:
        if self._stop is not None:
            self._stop.set()
        if self._thread is not None:
            self._thread.join(timeout=3)


def find_ribbon_button(window, label: str):
    wanted = normalise(label).lower()
    for control in window.descendants(control_type="Button"):
        try:
            if normalise(control.window_text()).lower() == wanted:
                return control
        except Exception:  # noqa: BLE001
            continue
    return None


def wait_for(predicate, timeout_s: float, what: str, poll_s: float = 1.0):
    deadline = time.monotonic() + timeout_s
    while time.monotonic() < deadline:
        try:
            found = predicate()
        except Exception:  # noqa: BLE001
            found = None
        if found:
            return found
        time.sleep(poll_s)
    raise ISpringPublishingError(
        f"timed out after {timeout_s}s waiting for {what}",
        user_message=USER_PUBLISH_FAILED,
    )


def open_publish_dialog(window):
    tab = window.child_window(title=RIBBON_TAB, control_type="TabItem")
    if not tab.exists():
        raise ISpringPublishingError(
            f"ribbon tab {RIBBON_TAB!r} not found — is the add-in loaded?",
            user_message=USER_PUBLISH_FAILED,
        )
    activate(tab, f"tab {RIBBON_TAB!r}")
    time.sleep(1)

    button = find_ribbon_button(window, "Publish")
    if button is None:
        raise ISpringPublishingError(
            "Publish button not found on the iSpring ribbon",
            user_message=USER_PUBLISH_FAILED,
        )
    activate(button, "ribbon Publish")

    dialog = wait_for(
        lambda: window.child_window(title=PUBLISH_DIALOG, control_type="Window"),
        timeout_s=60,
        what="the publish dialog",
    )
    dialog.wait("exists visible", timeout=60)
    _note("publish dialog open")
    return dialog


def destination_header(dialog) -> str:
    try:
        return normalise(
            dialog.child_window(auto_id=ID_DESTINATION_HEADER).window_text()
        )
    except Exception:  # noqa: BLE001
        return ""


def choose_cloud_destination(dialog) -> None:
    """The destination buttons on the left are unnamed, so go by the header."""
    header = destination_header(dialog)
    if "cloud" in header.lower():
        _note("destination already iSpring Cloud", header=header)
        return
    buttons = [
        control
        for control in dialog.descendants(control_type="Button")
        if not normalise(control.window_text())
    ]
    for index, button in enumerate(buttons):
        try:
            button.click_input()
        except Exception:  # noqa: BLE001
            continue
        time.sleep(1.5)
        header = destination_header(dialog)
        if "cloud" in header.lower():
            _note("selected iSpring Cloud", left_button=index + 1)
            return
    raise ISpringPublishingError(
        f"could not switch to iSpring Cloud (header still {header!r})",
        user_message=USER_PUBLISH_FAILED,
    )


def scroll_into_view(control, picker):
    """Get a project row on screen: ask nicely first, then scroll the list."""
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
                _note("scrolled project row into view", how=attempt)
                return rect
        except Exception:  # noqa: BLE001
            continue

    surface = None
    for pane in picker.descendants(control_type="Pane"):
        pane_rect = visible_rect(pane)
        if pane_rect is not None and pane_rect.height() > 200:
            surface = pane
            break
    surface = surface or picker
    for turn in range(80):
        try:
            surface.wheel_mouse_input(wheel_dist=-3)
        except Exception as exc:  # noqa: BLE001
            raise ISpringPublishingError(
                f"could not scroll the project list: {exc}",
                user_message=USER_PUBLISH_FAILED,
            ) from exc
        time.sleep(0.25)
        rect = visible_rect(control)
        if rect is not None:
            _note("scrolled project row into view", wheel_turns=turn + 1)
            return rect
    raise ISpringPublishingError(
        "scrolled to the end of the project list and the row never appeared",
        user_message=USER_PUBLISH_FAILED,
    )


def expand_branch(picker, label: str) -> bool:
    for control in find_named(picker, label):
        if visible_rect(control) is None:
            try:
                scroll_into_view(control, picker)
            except ISpringPublishingError:
                continue
        try:
            control.double_click_input()
            _note("expanded project branch", branch=label)
            time.sleep(2)
            return True
        except Exception:  # noqa: BLE001
            continue
    return False


def pick_project(dialog, institution: str, parent_folder: str) -> None:
    browse = dialog.child_window(auto_id=ID_BROWSE, control_type="Button")
    if not browse.exists():
        raise ISpringPublishingError(
            "Browse button not found on the publish dialog",
            user_message=USER_PUBLISH_FAILED,
        )
    activate(browse, "Browse")

    picker = wait_for(
        lambda: dialog.child_window(title=PROJECT_DIALOG, control_type="Window"),
        timeout_s=60,
        what="the project picker",
    )
    picker.wait("exists visible", timeout=60)

    matches = find_named(picker, institution)
    if not matches and parent_folder:
        expand_branch(picker, parent_folder)
        matches = find_named(picker, institution)
    if not matches:
        near = find_named(picker, institution, exact=False)
        hint = ", ".join(normalise(c.window_text()) for c in near[:5]) or "nothing similar"
        raise ISpringPublishingError(
            f"project {institution!r} is not in the tree (closest: {hint})",
            user_message=(
                f"No iSpring Cloud project named {institution!r} was found."
            ),
        )

    scroll_into_view(matches[0], picker)
    try:
        matches[0].click_input()
    except Exception as exc:  # noqa: BLE001
        raise ISpringPublishingError(
            f"could not click the project row: {exc}",
            user_message=USER_PUBLISH_FAILED,
        ) from exc

    select = picker.child_window(auto_id=ID_OK, control_type="Button")
    if not select.exists():
        raise ISpringPublishingError(
            "Select button not found in the project picker",
            user_message=USER_PUBLISH_FAILED,
        )
    activate(select, "Select")
    wait_for(lambda: not picker.exists(), timeout_s=30, what="the picker to close")
    _note("project chosen", project=institution)


def wait_for_completion(window, timeout_s: float) -> None:
    """Wait for the finished window.

    The progress window's status text stops updating and still reads
    "Uploading the presentation" after the upload is done, so it cannot be the
    finish signal. The "Publishing is complete!" window is.
    """

    def done():
        for control in by_auto_id(window, ID_DONE_TEXT):
            if normalise(control.window_text()):
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
            _note("publishing complete", message=normalise(finished.window_text()))
            for control in by_auto_id(window, ID_DONE_DETAIL):
                detail = normalise(control.window_text())
                if detail:
                    _note("publish detail", detail=detail)
            return
        current = status()
        if current and current != last_status:
            last_status = current
            _note("publish progress", operation=current)
        time.sleep(2)
    raise ISpringPublishingError(
        f"publish did not finish within {timeout_s}s (last status {last_status!r})",
        user_message="Publishing to iSpring Cloud timed out.",
    )


def click_manage_content(window) -> bool:
    """Press Manage Content, which opens the material in the browser."""
    buttons = by_auto_id(window, ID_DONE_MANAGE)
    if not buttons:
        _warn("no Manage Content button on the finish window")
        return False
    try:
        activate(buttons[0], "Manage Content")
        return True
    except ISpringPublishingError as exc:
        _warn("could not press Manage Content", error=str(exc))
        return False


def close_ispring_windows(window) -> None:
    """Close the finish and progress windows left over from the publish."""
    for child in window.descendants(control_type="Window"):
        title = normalise(child.window_text())
        if title != DONE_WINDOW and not title.startswith(PROGRESS_PREFIX):
            continue
        try:
            child.close()
            _note("closed iSpring window", window=title)
            continue
        except Exception:  # noqa: BLE001
            pass
        try:
            child.child_window(title="Close", control_type="Button").invoke()
            _note("closed iSpring window", window=title, how="Close button")
        except Exception:  # noqa: BLE001
            _warn("could not close iSpring window", window=title)


def close_powerpoint(app, pptx: Path) -> None:
    """Close the deck and quit PowerPoint, leaving nothing behind."""
    if app is None:
        return
    ensure_com()
    target = Path(pptx).expanduser().resolve()
    try:
        for presentation in list(app.Presentations):
            try:
                same = Path(str(presentation.FullName)).resolve() == target
            except Exception:  # noqa: BLE001
                same = False
            if same:
                try:
                    presentation.Saved = True  # do not prompt to save
                except Exception:  # noqa: BLE001
                    pass
                presentation.Close()
                _note("closed the deck", file=target.name)
    except Exception:  # noqa: BLE001
        _warn("could not close the deck", exc_info=True)
    try:
        app.Quit()
        _note("quit PowerPoint")
    except Exception:  # noqa: BLE001
        _warn("PowerPoint Quit failed; terminating", exc_info=True)
        from powerpoint.service import terminate_powerpoint_processes

        terminate_powerpoint_processes()


# ---------------------------------------------------------------------------
# Website half
# ---------------------------------------------------------------------------


def _frames(page):
    frames = [page]
    try:
        for frame in page.frames:
            if frame not in frames:
                frames.append(frame)
    except Exception:  # noqa: BLE001
        pass
    return frames


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


def safe_click(locator, what: str = "") -> bool:
    """Click something the page has covered with a styled overlay.

    iSpring's controls are real inputs under decorative divs, so a normal
    click is refused ("subtree intercepts pointer events"). Force skips that
    check; dispatching the event needs no coordinates at all.
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
            _note("clicked in browser", control=what or "control", method=name)
            return True
        except Exception:  # noqa: BLE001
            continue
    return False


def popup_root(page):
    """The open popup.

    Everything in a dialog must be looked for inside this element: searching
    the whole page matches text in the library list behind the popup, and the
    clicks then land there instead.
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
    root = popup_root(page)
    return root if root is not None else page


def click_by_role(page, pattern, roles=("button", "menuitem", "link", "tab"), what=""):
    for frame in _frames(page):
        for role in roles:
            found = visible_or_none(frame.get_by_role(role, name=pattern))
            if found is not None and safe_click(found, what or pattern.pattern):
                return True
    for frame in _frames(page):
        found = visible_or_none(frame.get_by_text(pattern))
        if found is not None and safe_click(found, what or pattern.pattern):
            return True
    return False


def scroll_hunt(page, name: str, tries: int = 40):
    """Find a row by name, scrolling the list until it appears.

    The library is a long list, not a search box, so a row further down is
    only reachable by scrolling and looking again.
    """
    pattern = re.compile(re.escape(name), re.I)
    for attempt in range(tries):
        for frame in _frames(page):
            found = visible_or_none(frame.get_by_text(pattern))
            if found is not None:
                try:
                    found.scroll_into_view_if_needed(timeout=3000)
                except Exception:  # noqa: BLE001
                    pass
                _note("found row in the library", row=name, scrolls=attempt)
                return found
        try:
            size = page.viewport_size or {"width": 1200, "height": 800}
            page.mouse.move(size["width"] / 2, size["height"] / 2)
            page.mouse.wheel(0, 700)
        except Exception:  # noqa: BLE001
            break
        page.wait_for_timeout(400)
    return None


def enter_folder(page, folder: str) -> bool:
    row = scroll_hunt(page, folder)
    if row is None:
        return False
    try:
        row.dblclick(timeout=4000)
        page.wait_for_timeout(2500)
        _note("opened folder", folder=folder)
        return True
    except Exception:  # noqa: BLE001
        pass
    if safe_click(row, f"folder {folder}"):
        page.wait_for_timeout(2500)
        return True
    return False


def open_row_menu(page, material: str) -> bool:
    """Open the three-dot menu on the material's row."""
    row = scroll_hunt(page, material)
    if row is None:
        return False
    try:
        row.hover()
        page.wait_for_timeout(500)
    except Exception:  # noqa: BLE001
        pass
    try:
        box = row.bounding_box() or {}
    except Exception:  # noqa: BLE001
        box = {}
    top = box.get("y", 0)
    bottom = top + box.get("height", 0)

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
        return False
    # The three dots are the right-most control on the row.
    candidates.sort(key=lambda pair: pair[0])
    return safe_click(candidates[-1][1], "three-dot menu")


def is_public(page) -> bool:
    return visible_or_none(scope(page).get_by_text(NOT_PUBLIC_RE)) is None


def toggle_candidates(page):
    """The first switch in the popup.

    Only the first: the second one is "Restrict with password", which must
    not be touched.
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
            if locator.count():
                found.append(locator.first)
        except Exception:  # noqa: BLE001
            continue
    return found


def ancestors(locator, levels: int = 4):
    chain = [locator]
    current = locator
    for _ in range(levels):
        try:
            current = current.locator("xpath=..")
            chain.append(current)
        except Exception:  # noqa: BLE001
            break
    return chain


def click_switch_beside_off(page) -> bool:
    """Click the visible switch, located by the word 'Off' printed beside it."""
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
    locator = root.get_by_text(OFF_RE)
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
        try:
            page.mouse.click(box["x"] + box["width"] + offset, y)
            page.wait_for_timeout(1500)
        except Exception:  # noqa: BLE001
            continue
        if is_public(page):
            _note("switched link sharing on", how=f"{offset}px right of Off")
            return True
    return False


def make_viewable_via_link(page) -> bool:
    """Turn on "Make viewable via link" and prove it went on.

    The embed code is shown even while the content is private, greyed out and
    pointing at something nobody can open, so this has to succeed first and is
    verified rather than assumed.
    """
    if is_public(page):
        _note("content already viewable via link")
        return True

    for attempt in range(1, 4):
        for candidate in toggle_candidates(page):
            for how, action in (
                ("space", lambda c=candidate: _press_space(c, page)),
                ("js click", lambda c=candidate: _js_click(c)),
            ):
                if not action():
                    continue
                page.wait_for_timeout(2000)
                if is_public(page):
                    _note("switched link sharing on", how=how)
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
                if not safe_click(target, f"link toggle wrapper {depth}"):
                    continue
                page.wait_for_timeout(2000)
                if is_public(page):
                    _note("switched link sharing on", how=f"wrapper {depth}")
                    return True
        _warn("link toggle still off", attempt=attempt)
        page.wait_for_timeout(1500)
    return False


def _press_space(locator, page) -> bool:
    """A real checkbox toggles on Space and fires its own change event."""
    try:
        locator.focus()
        page.keyboard.press("Space")
        return True
    except Exception:  # noqa: BLE001
        return False


def _js_click(locator) -> bool:
    """Call .click() inside the page, where nothing can intercept it."""
    try:
        locator.evaluate("el => el.click()")
        return True
    except Exception:  # noqa: BLE001
        return False


def read_embed_code(page) -> str:
    """Pull the iframe snippet out of the share popup."""
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
        for getter in (lambda f=frame: f.inner_text("body"), lambda f=frame: f.content()):
            try:
                blob = getter()
            except Exception:  # noqa: BLE001
                continue
            match = IFRAME_RE.search(blob or "")
            if match and "ispring" in match.group(0).lower():
                return match.group(0).strip()
    return ""


def close_popup(page) -> None:
    """Close the share popup so the browser is clean for the next job."""
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(800)
    except Exception:  # noqa: BLE001
        pass
    if popup_root(page) is None:
        _note("share popup closed")
        return
    for pattern in (re.compile(r"^\s*close\s*$", re.I), re.compile(r"^\s*(done|ok)\s*$", re.I)):
        if click_by_role(page, pattern, what="popup close"):
            page.wait_for_timeout(800)
            break
    if popup_root(page) is not None:
        _warn("share popup did not close")
    else:
        _note("share popup closed")


def fetch_embed(material: str, institution: str, cdp_url: str) -> str:
    """Drive the Chrome that is already running and already signed in.

    Nothing is launched and no profile is created, so a login done once by
    hand in that Chrome keeps working for every run after it.
    """
    from playwright.sync_api import sync_playwright

    with sync_playwright() as playwright:
        try:
            browser = playwright.chromium.connect_over_cdp(cdp_url)
        except Exception as exc:  # noqa: BLE001
            raise ISpringPublishingError(
                f"could not attach to Chrome at {cdp_url}: {exc}",
                user_message=USER_LINK_FAILED,
            ) from exc
        if not browser.contexts:
            raise ISpringPublishingError(
                "attached to Chrome but it has no windows open",
                user_message=USER_LINK_FAILED,
            )
        context = browser.contexts[0]
        pages = [p for p in context.pages if "ispring" in (p.url or "").lower()]
        if not pages:
            raise ISpringPublishingError(
                "no iSpring tab is open in that Chrome",
                user_message=USER_LINK_FAILED,
            )
        page = pages[-1]
        page.bring_to_front()
        page.wait_for_timeout(2000)

        if visible_or_none(page.get_by_text(re.compile(r"sign in|log in", re.I))) is not None:
            _warn("that Chrome may not be signed in to iSpring Cloud")

        if scroll_hunt(page, material, tries=4) is None and institution:
            if not enter_folder(page, institution):
                raise ISpringPublishingError(
                    f"could not open the folder {institution!r} in the library",
                    user_message=USER_LINK_FAILED,
                )

        if not open_row_menu(page, material):
            raise ISpringPublishingError(
                f"could not open the three-dot menu for {material!r}",
                user_message=USER_LINK_FAILED,
            )
        page.wait_for_timeout(1200)

        if not click_by_role(page, SHARE_RE, what="Share"):
            raise ISpringPublishingError(
                "no Share item in the row menu", user_message=USER_LINK_FAILED
            )
        page.wait_for_timeout(2000)

        if not make_viewable_via_link(page):
            raise ISpringPublishingError(
                "'Make viewable via link' could not be switched on; the embed "
                "code would point at content nobody can open",
                user_message=USER_LINK_FAILED,
            )
        page.wait_for_timeout(2000)

        embed = read_embed_code(page)
        if not embed:
            click_by_role(page, EMBED_RE, roles=("tab", "button", "link"), what="Embed")
            page.wait_for_timeout(1500)
            embed = read_embed_code(page)

        close_popup(page)
        if not embed:
            raise ISpringPublishingError(
                "the share popup showed no iframe code",
                user_message=USER_LINK_FAILED,
            )
        return embed


# ---------------------------------------------------------------------------
# The whole job
# ---------------------------------------------------------------------------


def publish_to_cloud(
    pptx: Path,
    *,
    institution: str,
    content_name: str,
    parent_folder: str = "PPT Migration",
    cdp_url: str = "http://127.0.0.1:9222",
    publish_timeout_s: float = 1800,
    close_powerpoint_after: bool = True,
    skip_open: bool = False,
) -> CloudPublishResult:
    """Publish one deck and return its embed URL.

    Everything opened along the way is closed again: the iSpring windows, the
    presentation, PowerPoint itself, and the share popup in the browser.

    ``skip_open`` is for callers that already opened the deck themselves — the
    job layer does, so that a damaged file goes through PowerPoint's repair
    first.
    """
    started = time.monotonic()
    app = None
    window = None
    try:
        if not skip_open:
            app = open_presentation(pptx)
        window = find_powerpoint_window()
        dismiss_nuisance_dialogs(window)
        answer_repair_prompt()

        dialog = open_publish_dialog(window)
        choose_cloud_destination(dialog)

        if content_name:
            field = dialog.child_window(auto_id=ID_CONTENT_NAME, control_type="Edit")
            if not field.exists():
                raise ISpringPublishingError(
                    "content name box not found on the publish dialog",
                    user_message=USER_PUBLISH_FAILED,
                )
            set_text(field, content_name, "content name")

        pick_project(dialog, institution, parent_folder)

        publish = dialog.child_window(auto_id=ID_OK, control_type="Button")
        if not publish.exists():
            raise ISpringPublishingError(
                "Publish button not found on the publish dialog",
                user_message=USER_PUBLISH_FAILED,
            )
        activate(publish, "Publish")
        wait_for_completion(window, publish_timeout_s)

        if not click_manage_content(window):
            raise ISpringPublishingError(
                "could not open the material in the browser",
                user_message=USER_LINK_FAILED,
            )
        time.sleep(10)  # let the browser open and settle

        embed = fetch_embed(content_name or institution, institution, cdp_url)
    finally:
        if window is not None:
            try:
                close_ispring_windows(window)
            except Exception:  # noqa: BLE001
                _warn("could not close the iSpring windows", exc_info=True)
        if close_powerpoint_after:
            try:
                close_powerpoint(app, Path(pptx))
            except Exception:  # noqa: BLE001
                _warn("could not close PowerPoint", exc_info=True)

    match = SRC_RE.search(embed)
    if not match:
        raise ISpringPublishingError(
            f"no src in the embed code: {embed[:200]}",
            user_message=USER_LINK_FAILED,
        )
    result = CloudPublishResult(
        iframe_url=match.group(1),
        embed_code=embed,
        elapsed_s=time.monotonic() - started,
    )
    _note("published to iSpring Cloud", iframe_url=result.iframe_url,
          seconds=round(result.elapsed_s))
    return result
