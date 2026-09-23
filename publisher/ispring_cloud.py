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
import base64
import os
import re
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Sequence
from urllib.parse import quote

from utils.exceptions import ISpringPublishingError, ProjectMissingError

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
# The tab is not always called that. An unlicensed machine shows
# "iSpring Free 11", and the version number moves with the release, so any
# iSpring tab counts and the Suite one is preferred when both exist.
ISPRING_TAB_RE = re.compile(r"^\s*ispring\b", re.I)
FREE_TAB_RE = re.compile(r"\bfree\b", re.I)

NUISANCE_TITLES = ("checking for updates", "update", "what's new")

# PowerPoint asks before repairing a damaged deck, and the dialog blocks the
# COM call that opened the file — so it is answered from another thread.
REPAIR_PROMPT_RE = re.compile(
    r"repair|recover|unreadable content|found a problem|damaged", re.I
)
REPAIR_BUTTONS = ("Repair", "Yes", "OK", "Open", "Continue")

# A retried queue row is published under the same job_id. iSpring then asks
# whether to replace the existing presentation — that must be Yes, not a
# second presentation with a new name.
OVERWRITE_PROMPT_RE = re.compile(
    r"already exists|overwrite|replace (it|the existing)|update the existing",
    re.I,
)
OVERWRITE_BUTTONS = ("Replace", "Overwrite", "Update", "Yes", "OK")

# --- website selectors ------------------------------------------------------
POPUP_SELECTORS = (
    '[data-at*="uikit-layer-popup"]',
    '[role="dialog"]',
    '[class*="uikit-layer"]',
    '[class*="popup"]',
    '[class*="modal"]',
)
# --- the library's own markup, read off harshit.ispring.com ----------------
# Every row carries data-at="selected=<bool>;id=row-<uuid>", the title is a
# link inside it, and the three dots only exist while the row is hovered.
ROW_SELECTOR = 'tr[data-at*="id=row-"]'
TITLE_SELECTOR = '[data-at="id=content-item-title"]'
ROW_MENU_BUTTON = '[data-at="id=table-shortcut-menu-button"]'
ROW_MENU_POPOVER = '[data-at="id=table-shortcut-menu-popover"]'
SEARCH_INPUT = 'input[data-at="id=search-global-input"]'
TABLE_BODY = '[data-at="id=table-body"]'

# The share popup, likewise read off the live page.
SHARE_MENU_ITEM = '[data-at="id=table-action-share"]'
SHARE_POPUP = '[data-at="id=sharing-popup"]'
PUBLIC_TOGGLE = '[data-at^="id=sharing-content-access-toggle"]'
# Never touched: switching this on puts a password on the content.
PASSWORD_TOGGLE = '[data-at^="id=sharing-content-password-access-toggle"]'
LINK_FIELD = '[data-at^="id=sharing-content-link-field"]'
EMBED_FIELD = '[data-at="id=sharing-embed-code-field"]'
COVER_BUTTON = '[data-at="id=customize-cover-button"]'
POPUP_CLOSE = '[data-at="id=popup-close-button"]'

IFRAME_RE = re.compile(r"<iframe\b[^>]*>(?:.*?</iframe>)?", re.I | re.S)
# The share popup shows the embed code as coloured text, not in a text box, and
# clips it with an ellipsis. So the URL is taken directly, and the iframe is
# rebuilt from it. The preview link carries the same id as the embed player.
EMBED_URL_RE = re.compile(r"https?://[^\s\"'<>]+/app/embed-player/[A-Za-z0-9\-]+", re.I)
PREVIEW_URL_RE = re.compile(r"(https?://[^\s\"'<>]+)/app/preview/([A-Za-z0-9\-]+)", re.I)
EMBED_TEMPLATE = (
    '<iframe src="{url}" width="560" height="315" frameborder="0" '
    'scrolling="auto" allowtransparency="true" allowfullscreen="1" '
    'style="border: none;"></iframe>'
)
SRC_RE = re.compile(r"""\bsrc\s*=\s*["']([^"']+)["']""", re.I)

# iSpring hands back a fixed-size iframe — width="1280" height="720". Inside
# the player that is a deck sitting in the corner of whatever box holds it,
# with the rest of the space empty and no way to fill it. The embed has to
# stretch instead, so the sizes are rewritten to 100% before it is stored.
IFRAME_OPEN_RE = re.compile(r"<iframe\b[^>]*>", re.I)
SIZE_ATTR_RE = re.compile(
    r"""\s(?:width|height)\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)""", re.I
)
MENU_BUTTON_RE = re.compile(r"more|menu|action|option|\.\.\.|…", re.I)
SHARE_RE = re.compile(r"^\s*share\b", re.I)
EMBED_RE = re.compile(r"embed", re.I)
PUBLIC_TOGGLE_RE = re.compile(r"viewable via link|public|share.*link", re.I)
NOT_PUBLIC_RE = re.compile(r"not publicly accessible", re.I)
OFF_RE = re.compile(r"^\s*off\s*$", re.I)

# After PowerPoint says publishing is complete, iSpring keeps uploading to
# the cloud for a few seconds. ISPRING_UPLOAD_WAIT overrides this.
UPLOAD_SETTLE_S = float(os.environ.get("ISPRING_UPLOAD_WAIT", "20"))
# How long to keep looking for the institution in the project picker
# while the tree loads its hundreds of folders.
PICKER_SEARCH_S = float(os.environ.get("ISPRING_PICKER_WAIT", "45"))
# Floor for how far to scroll the project list looking for a row.
SCROLL_TURNS_MIN = int(os.environ.get("ISPRING_SCROLL_TURNS", "250"))

USER_PUBLISH_FAILED = "The presentation could not be published to iSpring Cloud."
USER_LINK_FAILED = "The presentation was published but no share link could be read."


@dataclass(frozen=True)
class CloudPublishResult:
    iframe_url: str
    embed_code: str
    elapsed_s: float


def _note(event: str, **fields) -> None:
    # The parameter is named 'event', not 'message': a caller logging a field
    # called message would otherwise collide with it and raise TypeError.
    logger.info(event, extra={"stage": STAGE, **fields})


def _warn(event: str, **fields) -> None:
    logger.warning(event, extra={"stage": STAGE, **fields})


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


def find_powerpoint_window(deck: Path | str | None = None):
    """The PowerPoint window holding the deck.

    PowerPoint keeps more than one PPTFrameClass window around: a hidden
    frame it uses for automation, the start screen, and one per deck already
    open. Asking for the class alone raises "There are 2 elements that match"
    the moment a second one exists, which is most of the time on a machine
    somebody also works on.

    So every candidate is scored — visible first, then titled, then matching
    the deck's file name, then the largest — and the winner is looked up
    again by handle, which is unique. Passing ``deck`` is what makes the
    choice right rather than merely deterministic when several decks are
    open.
    """
    ensure_com()
    from pywinauto import Application  # type: ignore

    app = Application(backend="uia").connect(path="POWERPNT.EXE", timeout=30)
    wanted = Path(deck).stem.lower() if deck else ""

    def score(win):
        try:
            title = normalise(win.window_text())
            rect = win.rectangle()
            area = max(0, rect.width()) * max(0, rect.height())
            visible = bool(win.is_visible())
        except Exception:  # noqa: BLE001
            return None
        return (
            1 if visible else 0,
            1 if title else 0,
            1 if wanted and wanted in title.lower() else 0,
            area,
        ), title

    best = None
    best_score = None
    best_title = ""
    seen = 0
    deadline = time.monotonic() + 30
    while time.monotonic() < deadline:
        try:
            windows = app.windows(
                class_name="PPTFrameClass",
                top_level_only=True,
                visible_only=False,
                enabled_only=False,
            )
        except Exception:  # noqa: BLE001
            windows = []
        seen = len(windows)
        for win in windows:
            scored = score(win)
            if scored is None:
                continue
            value, title = scored
            if best_score is None or value > best_score:
                best, best_score, best_title = win, value, title
        # A visible window with a title is a real frame; anything less is
        # PowerPoint still starting up, so give it another second.
        if best_score is not None and best_score[0] and best_score[1]:
            break
        time.sleep(1)

    if best is None:
        raise ISpringPublishingError(
            "no PowerPoint window found",
            user_message=USER_PUBLISH_FAILED,
        )

    _note("found PowerPoint", title=best_title, windows=seen)
    return app.window(handle=best.handle)


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


def ribbon_tabs(window) -> list[str]:
    names = []
    try:
        for tab in window.descendants(control_type="TabItem"):
            name = normalise(tab.window_text())
            if name:
                names.append(name)
    except Exception:  # noqa: BLE001
        pass
    return names


def find_ispring_tab(window):
    """The iSpring tab on the ribbon, whatever it is called.

    The name carries both the edition and the version — "iSpring Suite 11",
    "iSpring Free 11" — so matching the exact string breaks on an unlicensed
    machine or after an upgrade. Suite wins when both are present.

    Returns ``(control, name)``, or ``(None, "")``.
    """
    found = []
    try:
        for tab in window.descendants(control_type="TabItem"):
            name = normalise(tab.window_text())
            if ISPRING_TAB_RE.match(name):
                found.append((tab, name))
    except Exception:  # noqa: BLE001
        return None, ""
    if not found:
        return None, ""
    found.sort(key=lambda pair: 0 if "suite" in pair[1].lower() else 1)
    return found[0]


def warn_if_free_edition(name: str) -> None:
    if name and FREE_TAB_RE.search(name):
        _warn(
            "PowerPoint is showing the iSpring Free ribbon, not Suite. The "
            "Suite licence is not active on this machine, and Free may not "
            "offer publishing to iSpring Cloud",
            tab=name,
        )


def has_addin_tab(window, timeout_s: float = 20, require_suite: bool = True) -> bool:
    """Is the Suite tab on the ribbon? Waits a little, never raises.

    A ribbon showing only "iSpring Free 11" counts as *not ready*: both
    add-ins are registered on this machine, and when Office disables the
    Suite one the Free one is left behind. That looks like iSpring is
    loaded while the Suite publish options are gone. Treating it as missing
    is what makes the automatic repair run.
    """
    deadline = time.monotonic() + timeout_s
    while True:
        tab, name = find_ispring_tab(window)
        if tab is not None:
            if not require_suite or not FREE_TAB_RE.search(name):
                _note("iSpring ribbon tab present", tab=name)
                return True
            warn_if_free_edition(name)
        if time.monotonic() >= deadline:
            return False
        dismiss_nuisance_dialogs(window)
        time.sleep(2)


def wait_for_addin(window, timeout_s: float = 90):
    """Wait for the iSpring tab to appear on the ribbon.

    The add-in loads a few seconds after PowerPoint starts, so a job that
    restarts PowerPoint arrives before the tab exists. Waiting beats failing.
    """
    deadline = time.monotonic() + timeout_s
    waited = False
    while time.monotonic() < deadline:
        tab, name = find_ispring_tab(window)
        if tab is not None:
            if waited:
                _note("iSpring ribbon tab appeared", tab=name)
            else:
                _note("iSpring ribbon tab found", tab=name)
            warn_if_free_edition(name)
            return tab
        waited = True
        dismiss_nuisance_dialogs(window)
        time.sleep(2)
    raise ISpringPublishingError(
        f"no iSpring tab appeared after {timeout_s:.0f}s; "
        f"tabs on the ribbon: {ribbon_tabs(window)}",
        user_message=(
            "The iSpring add-in did not load in PowerPoint. Check it is "
            "enabled in File > Options > Add-ins."
        ),
    )


def open_publish_dialog(window):
    tab = wait_for_addin(window)
    activate(tab, f"tab {normalise(tab.window_text())!r}")
    time.sleep(1)

    button = find_ribbon_button(window, "Publish")
    if button is None:
        raise ISpringPublishingError(
            "Publish button not found on the iSpring ribbon; tabs present: "
            f"{ribbon_tabs(window)}",
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
    # 80 turns was set when the tree held one parent. With both open it holds
    # 340 + 138 folders, and three lines a turn does not get to the bottom —
    # the row is there, the scroll just stops short, and the error says the
    # project was never found. Scale it to what is actually on screen.
    turns = max(SCROLL_TURNS_MIN, len(picker.descendants()) // 2)
    for turn in range(turns):
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
        f"scrolled {turns} turns and the row never appeared",
        user_message=USER_PUBLISH_FAILED,
    )


def resolve_parent(picker, label: str) -> str:
    """The row in the tree that is this parent, however it is written there.

    The configured name has to match the tree exactly, and a tree label is
    not always the plain project name — the library shows several projects
    with the owner appended, as in "Edmingle Owner (server@edmingle.com)".
    One character out and the branch is never opened, so every folder under
    it is invisible.

    Exact wins. Failing that, the shortest row containing the name: "PPT
    Migration" is a prefix of "PPT migration New", so longest-match or
    first-match would quietly open the wrong branch.
    """
    exact = find_named(picker, label)
    if exact:
        return normalise(exact[0].window_text())
    near = find_named(picker, label, exact=False)
    if not near:
        return ""
    return min((normalise(c.window_text()) for c in near), key=len)


def expand_branch(picker, label: str) -> str:
    """Open a parent branch. Returns what happened, for the log.

    One of: "open" (it was already), "expanded", "expanded by double click",
    "not in the tree", or "could not expand". It used to return a bare
    False for the last two, which meant a parent that was never opened —
    or never even there — looked exactly like one that opened fine, and
    every folder under it was invisible with nothing saying why.
    """
    found = find_named(picker, label)
    if not found:
        return "not in the tree"

    for control in found:
        if visible_rect(control) is None:
            try:
                scroll_into_view(control, picker)
            except ISpringPublishingError:
                continue
        # Ask the tree to expand. A double-click toggles, so on a branch that
        # is already open it would close it and hide the very rows being
        # looked for.
        try:
            expander = control.iface_expand_collapse
            if expander.CurrentExpandCollapseState == 1:  # already expanded
                return "open"
            expander.Expand()
            time.sleep(2)
            return "expanded"
        except Exception:  # noqa: BLE001
            pass
        try:
            control.double_click_input()
            time.sleep(2)
            return "expanded by double click"
        except Exception:  # noqa: BLE001
            continue
    return "could not expand"


def ancestor_label(control, labels: Sequence[str], max_up: int = 8) -> str:
    """Which of ``labels`` this row sits under, or "" when none of them does."""
    wanted = {normalise(label).lower(): label for label in labels if label}
    node = control
    for _ in range(max_up):
        try:
            node = node.parent()
        except Exception:  # noqa: BLE001
            return ""
        if node is None:
            return ""
        try:
            text = normalise(node.window_text()).lower()
        except Exception:  # noqa: BLE001
            continue
        if text in wanted:
            return wanted[text]
    return ""


def dismiss_window(win, what: str) -> None:
    """Close a dialog so the next attempt starts from a clean screen."""
    try:
        win.close()
        return
    except Exception:  # noqa: BLE001
        pass
    try:
        win.child_window(title="Cancel", control_type="Button").invoke()
    except Exception:  # noqa: BLE001
        _warn("could not close a dialog", window=what)


def choose_match(placed: Sequence[tuple], parents: Sequence[str], institution: str) -> list:
    """Narrow rows that share the institution's name down to the one to click.

    ``placed`` pairs each row with the parent it sits under, "" when none.
    The first parent listed wins a clash, so the same institution always
    lands in the same folder rather than wherever the tree happened to
    yield it first.
    """
    if not placed or not parents:
        return [control for _, control in placed]

    under = [(parent, control) for parent, control in placed if parent]
    if not under:
        # No row reports one of the parents as an ancestor. That is a flat
        # tree, not a wrong folder: keep the matches rather than throw away a
        # folder that is really there.
        _warn(
            "could not tell which parent the folder sits under",
            institution=institution,
        )
        return [control for _, control in placed]

    seen = sorted({parent for parent, _ in under})
    if len(seen) > 1:
        _warn(
            "this institution has a folder under both parents",
            institution=institution,
            parents=seen,
            using=parents[0],
        )
    order = {parent: index for index, parent in enumerate(parents)}
    under.sort(key=lambda pair: order[pair[0]])
    return [under[0][1]]


def pick_project(dialog, institution: str, parent_folders: Sequence[str]) -> None:
    """Select the institution's folder, looking only under ``parent_folders``.

    Scoped on purpose. Four institution names exist under both parents, and a
    search of the whole tree took whichever row the tree yielded first, which
    is not a choice anyone made. Here the first parent in the list wins and
    the clash is logged.

    A missing folder raises ProjectMissingError rather than the general
    publishing error: the caller creates the folder and tries again.
    """
    parents = [folder for folder in parent_folders if folder]
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

    # Expand, then keep looking until it turns up.
    #
    # The tree fills itself in after a branch opens, and these branches hold
    # 340 and 133 folders. Looking once, straight after expanding, reads a
    # tree that is still loading and finds nothing — which is indistinguishable
    # from the folder not existing, and is why a folder that was there all
    # along came back as missing. So the branches are opened and then the
    # search is repeated until the row appears or the time runs out.
    # What the configured parents are actually called in this tree.
    resolved = []
    for parent in parents:
        actual = resolve_parent(picker, parent)
        _note("parent folder", configured=parent, found_as=actual or "(not found)")
        if actual:
            resolved.append(actual)
    if resolved:
        parents = resolved

    matches = []
    deadline = time.monotonic() + PICKER_SEARCH_S
    rounds = 0
    while time.monotonic() < deadline:
        rounds += 1
        for parent in parents:
            outcome = expand_branch(picker, parent)
            # Every parent, every round. If one of them never opens, this is
            # the line that says so — and a parent that never opens is a
            # parent whose institutions cannot be found.
            _note(
                "project branch",
                branch=parent,
                outcome=outcome,
                round=rounds,
                rows=len(picker.descendants()),
            )
        matches = choose_match(
            [(ancestor_label(control, parents), control)
             for control in find_named(picker, institution)],
            parents,
            institution,
        )
        if matches:
            break
        time.sleep(2)

    if matches and rounds > 1:
        _note("the folder appeared once the tree finished loading",
              institution=institution, rounds=rounds)

    if not matches:
        near = find_named(picker, institution, exact=False)
        hint = ", ".join(normalise(c.window_text()) for c in near[:5]) or "nothing similar"
        # What the picker was actually holding. Without this the error says
        # only that the name was not there, which is the one thing already
        # known, and every diagnosis after it is guesswork.
        listed = [
            name for name in (
                normalise(c.window_text()) for c in picker.descendants()
            ) if name
        ]
        _warn(
            "the picker did not have this folder",
            institution=institution,
            rows=len(listed),
            sample=listed[:40],
        )
        dismiss_window(picker, "the project picker")
        dismiss_window(dialog, "the publish dialog")
        raise ProjectMissingError(
            f"{institution!r} has no folder under {parents} after "
            f"{PICKER_SEARCH_S:.0f}s and {rounds} looks "
            f"({len(listed)} rows in the picker, closest: {hint})",
            institution=institution,
            user_message=(
                f"No iSpring Cloud folder named {institution!r} was found."
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
            _note("publishing complete", text=normalise(finished.window_text()))
            for control in by_auto_id(window, ID_DONE_DETAIL):
                detail = normalise(control.window_text())
                if detail:
                    _note("publish detail", detail=detail)
            return
        current = status()
        if current and current != last_status:
            last_status = current
            _note("publish progress", operation=current)
        confirm_overwrite(window)
        time.sleep(2)
    raise ISpringPublishingError(
        f"publish did not finish within {timeout_s}s (last status {last_status!r})",
        user_message="Publishing to iSpring Cloud timed out.",
    )


def is_overwrite_prompt(text: str) -> bool:
    return bool(OVERWRITE_PROMPT_RE.search(normalise(text)))


def preferred_overwrite_button(labels: list[str]) -> str | None:
    indexed = [(normalise(label).lower(), label) for label in labels]
    for wanted in OVERWRITE_BUTTONS:
        key = wanted.lower()
        for normalised, original in indexed:
            if normalised == key or normalised.startswith(key):
                return original
    return None


def confirm_overwrite(window) -> bool:
    """Click through iSpring's 'this name already exists — replace?' dialog.

    Retry jobs reuse job_id as the presentation name. Creating a second
    presentation would leave the failed copy in place.
    """
    try:
        dialogs = list(window.descendants(control_type="Window"))
        dialogs.append(window)
    except Exception:  # noqa: BLE001
        return False
    answered = False
    for dialog in dialogs:
        try:
            texts = [normalise(dialog.window_text())]
            for child in dialog.descendants(control_type="Text")[:20]:
                texts.append(normalise(child.window_text()))
        except Exception:  # noqa: BLE001
            continue
        haystack = " ".join(part for part in texts if part)
        if not is_overwrite_prompt(haystack):
            continue
        try:
            labels = [
                btn.window_text()
                for btn in dialog.descendants(control_type="Button")
                if btn.window_text()
            ]
        except Exception:  # noqa: BLE001
            continue
        wanted = preferred_overwrite_button(labels)
        if not wanted:
            continue
        try:
            button = dialog.child_window(title=wanted, control_type="Button")
            if not button.exists():
                continue
            button.invoke()
            _note("overwrote existing presentation", button=wanted)
            answered = True
        except Exception:  # noqa: BLE001
            continue
    return answered


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
    """The page, then any iframes in it.

    The page's own main frame is left out: it is the same document as the
    page, and including it searches — and dumps — everything twice.
    """
    frames = [page]
    try:
        main = page.main_frame
    except Exception:  # noqa: BLE001
        main = None
    try:
        for frame in page.frames:
            if frame is not main and frame not in frames:
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


# ---------------------------------------------------------------------------
# Browser dump
#
# The desktop half has scripts/probe_ispring_ui.py to print a window's whole
# control tree when something is not where it was expected. This is the same
# thing for the website: it writes out every visible element, every open
# popup and the page's text, so a failed run can be read afterwards instead
# of guessed at.
# ---------------------------------------------------------------------------

DUMP_JS = r"""
() => {
  const MAX_ROWS = 3000;
  const rows = [];
  const isVisible = (el) => {
    try {
      const r = el.getBoundingClientRect();
      if (r.width < 1 || r.height < 1) return false;
      const s = window.getComputedStyle(el);
      if (s.visibility === 'hidden' || s.display === 'none') return false;
      if (parseFloat(s.opacity || '1') < 0.05) return false;
      return true;
    } catch (e) { return false; }
  };
  const ownText = (el) => {
    let out = '';
    for (const node of el.childNodes) {
      if (node.nodeType === 3) out += node.nodeValue;
    }
    return out.replace(/\s+/g, ' ').trim();
  };
  const interesting = (el) => {
    const tag = el.tagName.toLowerCase();
    if (['a','button','input','textarea','select','summary','label','iframe'].includes(tag)) return true;
    if (el.hasAttribute('role') || el.hasAttribute('aria-label') || el.hasAttribute('data-at')) return true;
    if (el.getAttribute('contenteditable') === 'true') return true;
    return ownText(el).length > 0;
  };
  const describe = (el, depth) => {
    const tag = el.tagName.toLowerCase();
    let head = tag;
    if (el.id) head += '#' + el.id;
    const cls = (el.getAttribute('class') || '').trim().split(/\s+/).filter(Boolean).slice(0, 3);
    if (cls.length) head += '.' + cls.join('.');
    const attrs = [];
    const wanted = ['role','data-at','aria-label','aria-checked','aria-expanded',
                    'aria-hidden','title','placeholder','href','type','name',
                    'disabled','tabindex'];
    for (const name of wanted) {
      if (el.hasAttribute(name)) {
        let v = el.getAttribute(name) || '';
        if (v.length > 140) v = v.slice(0, 140) + '…';
        attrs.push(name + '=' + JSON.stringify(v));
      }
    }
    if ((tag === 'input' || tag === 'textarea') && el.value) {
      attrs.push('value=' + JSON.stringify(String(el.value).slice(0, 400)));
    }
    if (tag === 'input' && el.type === 'checkbox') attrs.push('checked=' + el.checked);
    const r = el.getBoundingClientRect();
    const rect = '@' + Math.round(r.x) + ',' + Math.round(r.y) +
                 ' ' + Math.round(r.width) + 'x' + Math.round(r.height);
    let text = ownText(el);
    if (text.length > 180) text = text.slice(0, 180) + '…';
    return '  '.repeat(Math.min(depth, 18)) + head +
      (attrs.length ? ' [' + attrs.join(' ') + ']' : '') + ' ' + rect +
      (text ? '  "' + text + '"' : '');
  };
  const walk = (el, depth) => {
    if (rows.length > MAX_ROWS) return;
    for (const child of el.children) {
      const tag = child.tagName.toLowerCase();
      if (tag === 'script' || tag === 'style' || tag === 'noscript') continue;
      if (!isVisible(child)) continue;
      if (interesting(child)) {
        rows.push(describe(child, depth));
        walk(child, depth + 1);
      } else {
        walk(child, depth);
      }
    }
  };
  if (document.body) walk(document.body, 0);
  const values = [];
  for (const el of document.querySelectorAll('input, textarea')) {
    if (el.value) values.push((el.getAttribute('aria-label') || el.getAttribute('name') ||
                               el.type || 'field') + ' = ' + String(el.value).slice(0, 500));
  }
  return {
    url: location.href,
    title: document.title,
    size: window.innerWidth + 'x' + window.innerHeight,
    tree: rows,
    values: values,
    text: (document.body ? document.body.textContent || '' : '').replace(/\s+/g, ' ').trim().slice(0, 20000),
  };
}
"""

URL_RE = re.compile(r"https?://[^\s\"'<>\\]+", re.I)


def dump_always() -> bool:
    """ISPRING_DUMP=1 dumps the browser on every run, not only on failure."""
    return os.environ.get("ISPRING_DUMP", "").strip().lower() in {"1", "true", "yes", "on"}


def dump_dir() -> Path:
    """Where browser dumps are written. ISPRING_DUMP_DIR overrides it."""
    raw = os.environ.get("ISPRING_DUMP_DIR", "")
    folder = Path(raw) if raw else Path.cwd() / "test-artifacts"
    try:
        folder.mkdir(parents=True, exist_ok=True)
    except Exception:  # noqa: BLE001
        folder = Path(os.environ.get("TEMP", ".")).resolve()
    return folder


def _popup_sections(page) -> list[str]:
    """Every open popup, with its markup — the embed code lives in one."""
    sections: list[str] = []
    for frame in _frames(page):
        for selector in POPUP_SELECTORS:
            try:
                locator = frame.locator(selector)
                count = locator.count()
            except Exception:  # noqa: BLE001
                continue
            for index in range(min(count, 6)):
                item = locator.nth(index)
                try:
                    if not item.is_visible():
                        continue
                    text = item.evaluate("el => el.textContent || ''")
                    markup = item.evaluate("el => el.outerHTML || ''")
                except Exception:  # noqa: BLE001
                    continue
                flat = re.sub(r"[ \t]+", " ", text or "").strip()[:4000]
                sections.append(
                    f"--- popup {selector} [{index}] ---\n"
                    f"text: {flat}\n"
                    f"html: {markup[:8000]}"
                )
    return sections


def dump_page(page, label: str = "page") -> str:
    """Write the whole state of the browser to a text file.

    Returns the file path, or "" when nothing could be written. Never raises:
    a dump is diagnostics, and failing to take one must not replace the error
    that prompted it.
    """
    stamp = time.strftime("%Y%m%d-%H%M%S")
    safe = re.sub(r"[^A-Za-z0-9._-]+", "-", label).strip("-") or "page"
    target = dump_dir() / f"ispring-browser-{safe}-{stamp}.txt"
    lines = [
        "=" * 78,
        f"iSpring browser dump: {label}",
        f"taken: {time.strftime('%Y-%m-%d %H:%M:%S')}",
        "=" * 78,
    ]
    try:
        lines.append(f"page url:   {page.url}")
    except Exception:  # noqa: BLE001
        pass

    try:
        popups = _popup_sections(page)
    except Exception as exc:  # noqa: BLE001
        popups = [f"(popup scan failed: {exc})"]
    lines.append("")
    lines.append(f"OPEN POPUPS: {len(popups)}")
    lines.extend(popups)

    urls: set[str] = set()
    for index, frame in enumerate(_frames(page)):
        try:
            data = frame.evaluate(DUMP_JS)
        except Exception as exc:  # noqa: BLE001
            lines.append(f"\n--- frame {index}: could not be read ({exc}) ---")
            continue
        lines.append("")
        lines.append("=" * 78)
        lines.append(f"FRAME {index}: {data.get('title', '')}")
        lines.append(f"url:  {data.get('url', '')}")
        lines.append(f"size: {data.get('size', '')}")
        lines.append("=" * 78)
        values = data.get("values") or []
        if values:
            lines.append("")
            lines.append("-- field values --")
            lines.extend(f"  {value}" for value in values)
        lines.append("")
        lines.append("-- visible elements --")
        lines.extend(data.get("tree") or [])
        text = data.get("text") or ""
        lines.append("")
        lines.append("-- page text --")
        lines.append(text)
        for match in URL_RE.finditer(text):
            urls.add(match.group(0))

    interesting = sorted(
        url for url in urls if re.search(r"embed|preview|player|share", url, re.I)
    )
    if interesting:
        lines.append("")
        lines.append("-- urls worth a look --")
        lines.extend(f"  {url}" for url in interesting)

    try:
        target.write_text("\n".join(lines), encoding="utf-8", errors="replace")
    except Exception as exc:  # noqa: BLE001
        _warn("could not write the browser dump", error=str(exc))
        return ""
    try:
        page.screenshot(path=str(target.with_suffix(".png")), full_page=False)
    except Exception:  # noqa: BLE001
        pass
    _note("wrote a browser dump", file=str(target), label=label)
    return str(target)


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
    # Start from the top of the list. A previous hunt may have left the page
    # scrolled half way down, and the newest material sits at the top — which
    # is how a row that is plainly on screen gets scrolled straight past.
    try:
        page.evaluate("() => window.scrollTo(0, 0)")
        for selector in ("[role=grid]", "[role=table]", "main", "[class*=scroll]"):
            page.evaluate(
                "sel => document.querySelectorAll(sel)"
                ".forEach(el => { if (el.scrollTop) el.scrollTop = 0; })",
                selector,
            )
        page.wait_for_timeout(400)
    except Exception:  # noqa: BLE001
        pass
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


def wait_for_rows(page, timeout_s: float = 25) -> int:
    """Wait until the content table has rows in it.

    The list is drawn a few seconds after the page loads, so looking straight
    after a navigation finds an empty table and reports the material missing.
    """
    deadline = time.monotonic() + timeout_s
    while True:
        try:
            count = page.locator(ROW_SELECTOR).count()
        except Exception:  # noqa: BLE001
            count = 0
        if count:
            return count
        if time.monotonic() >= deadline:
            _warn("the content table is still empty", waited_s=round(timeout_s))
            return 0
        page.wait_for_timeout(1000)


def row_titles(page, limit: int = 20) -> list[str]:
    """The titles of the rows on screen, in order."""
    titles: list[str] = []
    try:
        locator = page.locator(f"{ROW_SELECTOR} {TITLE_SELECTOR}")
        count = min(locator.count(), limit)
    except Exception:  # noqa: BLE001
        return titles
    for index in range(count):
        try:
            titles.append((locator.nth(index).inner_text() or "").strip())
        except Exception:  # noqa: BLE001
            continue
    return titles


ROW_SCAN_JS = r"""
() => [...document.querySelectorAll('tr[data-at*="id=row-"]')].map((row, index) => ({
  index,
  id: (row.getAttribute('data-at') || '').replace(/^.*id=row-/, ''),
  title: (row.querySelector('[data-at="id=content-item-title"]')?.innerText || '').trim(),
  cells: [...row.children].map(c => (c.innerText || '').replace(/\s+/g, ' ').trim()).filter(Boolean),
}))
"""

# "Sep 17, 2026, 3:18 PM" — the modified column.
ROW_DATE_RE = re.compile(r"[A-Z][a-z]{2}\s+\d{1,2},\s+\d{4},\s+\d{1,2}:\d{2}\s*[AP]M")


def scan_rows(page) -> list[dict]:
    """Every row on screen: its id, title and the text of each cell."""
    for frame in _frames(page):
        try:
            rows = frame.evaluate(ROW_SCAN_JS)
        except Exception:  # noqa: BLE001
            continue
        if rows:
            return rows
    return []


def row_modified(row: dict):
    """When the row says it was last modified, for choosing between twins."""
    from datetime import datetime

    for cell in row.get("cells", []):
        match = ROW_DATE_RE.search(cell)
        if not match:
            continue
        text = re.sub(r"\s+", " ", match.group(0))
        for pattern in ("%b %d, %Y, %I:%M %p", "%b %d, %Y %I:%M %p"):
            try:
                return datetime.strptime(text, pattern)
            except ValueError:
                continue
    return None


def row_folder(row: dict) -> str:
    """The project/folder cell, which iSpring writes as "Parent / Child"."""
    for cell in row.get("cells", []):
        if "/" in cell:
            return cell
    cells = row.get("cells", [])
    return cells[2] if len(cells) > 2 else ""


def check_folder(row: dict, institution: str) -> None:
    """Say so when the deck did not land in the institution's own project.

    iSpring drops the content in the parent folder when the institution has
    no project of its own, and says nothing about it. The iframe still works,
    so this is a warning, not a failure.
    """
    if not institution:
        return
    folder = row_folder(row)
    if folder and institution.casefold() not in folder.casefold():
        _warn(
            "the material is not in the institution's own project",
            folder=folder,
            institution=institution,
        )
    elif folder:
        _note("material is in the right project", folder=folder)


def find_row(page, name: str, scrolls: int = 12, institution: str = ""):
    """The row whose title is this material, by the library's own markup.

    Exact title first: a search for "FA 4" also lists "20208952-FA_4
    [Repaired]", and opening the wrong one gives the wrong iframe. When two
    rows carry the same title — the same deck published twice — the newest
    one wins, because that is the one just published.
    """
    wanted = normalise(name).casefold()
    wait_for_rows(page)
    for attempt in range(max(1, scrolls)):
        rows = scan_rows(page)
        exact = [r for r in rows if normalise(r.get("title", "")).casefold() == wanted]
        if exact:
            if len(exact) > 1:
                exact.sort(
                    key=lambda r: (row_modified(r) is not None, row_modified(r)),
                    reverse=True,
                )
                _warn(
                    "several materials share this title; taking the newest",
                    title=name,
                    copies=len(exact),
                    folders=[row_folder(r) for r in exact][:4],
                )
            chosen = exact[0]
            check_folder(chosen, institution)
            _note(
                "found the row",
                title=chosen.get("title"),
                folder=row_folder(chosen),
                scrolls=attempt,
            )
            return row_locator(page, chosen)

        loose = [r for r in rows if wanted and wanted in normalise(r.get("title", "")).casefold()]
        if loose:
            _note("no exact title; using the closest row", title=loose[0].get("title"))
            check_folder(loose[0], institution)
            return row_locator(page, loose[0])

        try:
            size = page.viewport_size or {"width": 1200, "height": 800}
            page.mouse.move(size["width"] / 2, size["height"] / 2)
            page.mouse.wheel(0, 700)
            page.wait_for_timeout(500)
        except Exception:  # noqa: BLE001
            break
    return None


def row_locator(page, row: dict):
    """A handle on the row this scan described, by its own id."""
    locator = None
    identifier = row.get("id") or ""
    if identifier:
        locator = page.locator(f'tr[data-at*="id=row-{identifier}"]').first
    if locator is None:
        locator = page.locator(ROW_SELECTOR).nth(int(row.get("index", 0)))
    try:
        locator.scroll_into_view_if_needed(timeout=3000)
    except Exception:  # noqa: BLE001
        pass
    return locator


def open_row_menu_for(page, row) -> bool:
    """Hover the row and press its three dots.

    The button is not in the page until the row is hovered, and only a real
    mouse move triggers that — a synthetic event does nothing.
    """
    try:
        row.hover(timeout=5000)
        page.wait_for_timeout(600)
    except Exception as exc:  # noqa: BLE001
        _warn("could not hover the row", error=str(exc))
        return False
    button = row.locator(ROW_MENU_BUTTON)
    try:
        button.wait_for(state="visible", timeout=5000)
    except Exception:  # noqa: BLE001
        _warn("the row's menu button did not appear on hover")
        return False
    if not safe_click(button.first, "three dots"):
        return False
    try:
        page.locator(ROW_MENU_POPOVER).first.wait_for(state="visible", timeout=6000)
    except Exception:  # noqa: BLE001
        _warn("the row menu did not open")
        return False
    return True


def visible_row_labels(page, limit: int = 25) -> list[str]:
    """The names of the rows on screen right now.

    Printed into the log when the material cannot be found: nine times out
    of ten it shows the name is spelled differently from what was asked for.
    """
    script = """
    () => {
      const out = [];
      const seen = new Set();
      const rows = document.querySelectorAll(
        '[role=row], [role=gridcell], [role=listitem], tr, a[href*="/app/"]');
      for (const row of rows) {
        const text = (row.textContent || '').replace(/\\s+/g, ' ').trim();
        if (!text || text.length > 120 || seen.has(text)) continue;
        const r = row.getBoundingClientRect();
        if (r.width < 1 || r.height < 1) continue;
        seen.add(text);
        out.push(text);
      }
      return out;
    }
    """
    labels: list[str] = []
    for frame in _frames(page):
        try:
            labels.extend(frame.evaluate(script))
        except Exception:  # noqa: BLE001
            continue
    return labels[:limit]


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


def row_element(page, text_node):
    """The whole list row that holds this text.

    The name is a small span in the middle of a wide row, and the three dots
    sit at the row's right edge. Measuring the span instead of the row is how
    the menu of a neighbouring row gets opened, so climb until the element is
    as wide as the list.
    """
    try:
        width = (page.viewport_size or {}).get("width") or 1200
    except Exception:  # noqa: BLE001
        width = 1200
    current = text_node
    for _ in range(6):
        try:
            box = current.bounding_box() or {}
        except Exception:  # noqa: BLE001
            return text_node
        if box.get("width", 0) >= width * 0.35:
            return current
        try:
            parent = current.locator("xpath=..")
            if not parent.count():
                return current
            current = parent.first
        except Exception:  # noqa: BLE001
            return current
    return current


def open_row_menu(page, material: str) -> bool:
    """Open the three-dot menu on the material's row."""
    found = scroll_hunt(page, material)
    if found is None:
        return False
    row = row_element(page, found)
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
                # Inside the row horizontally too, or a toolbar button that
                # happens to sit at the same height gets clicked instead.
                if box:
                    left = box.get("x", 0)
                    right = left + box.get("width", 0)
                    if not (left - 4 <= item_box.get("x", 0) <= right + 4):
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
    state = toggle_state(page)
    if state is not None:
        return state
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


def toggle_state(page) -> bool | None:
    """Is link sharing on? Read from the site's own state attribute.

    The toggle carries data-at="id=sharing-content-access-toggle;state=true"
    or ";state=false", which is a far better answer than guessing from the
    words on screen. None means the toggle was not found.
    """
    for frame in _frames(page):
        try:
            locator = frame.locator(PUBLIC_TOGGLE)
            count = locator.count()
        except Exception:  # noqa: BLE001
            continue
        for index in range(min(count, 3)):
            try:
                value = locator.nth(index).get_attribute("data-at") or ""
            except Exception:  # noqa: BLE001
                continue
            if "state=true" in value:
                return True
            if "state=false" in value:
                return False
    return None


def switch_public_toggle(page) -> bool:
    """Click the link-sharing switch itself, by its data-at.

    Deliberately never the second switch on the popup: that one is
    "Restrict with password".
    """
    for frame in _frames(page):
        try:
            holder = visible_or_none(frame.locator(PUBLIC_TOGGLE))
        except Exception:  # noqa: BLE001
            continue
        if holder is None:
            continue
        for what, getter in (
            ("checkbox", lambda h=holder: h.locator("input[type=checkbox]").first),
            ("switch", lambda h=holder: h),
        ):
            try:
                target = getter()
            except Exception:  # noqa: BLE001
                continue
            if _js_click(target) or safe_click(target, f"link toggle ({what})"):
                page.wait_for_timeout(2000)
                if toggle_state(page):
                    return True
    return False


def make_viewable_via_link(page) -> bool:
    """Turn on "Make viewable via link" and prove it went on.

    The embed code is shown even while the content is private, greyed out and
    pointing at something nobody can open, so this has to succeed first and is
    verified rather than assumed.
    """
    state = toggle_state(page)
    if state is True:
        _note("content already viewable via link")
        return True
    if state is False and switch_public_toggle(page):
        _note("switched link sharing on")
        return True

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
    """Get the embed iframe out of the share popup.

    Four ways, in order of how much they can be trusted:
    the text box if there is one; the popup's text content (which holds the
    whole URL even though the display clips it); the preview link, which
    carries the same id as the embed player; and finally the Copy button.
    """
    # The site puts the whole iframe in one element. Take it from there.
    for frame in _frames(page):
        try:
            field = visible_or_none(frame.locator(EMBED_FIELD))
            if field is not None:
                text = (field.evaluate("el => el.textContent || ''") or "").strip()
                if "<iframe" in text.lower():
                    _note("read the embed code from the share popup")
                    return " ".join(text.split())
        except Exception:  # noqa: BLE001
            continue

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

    # textContent, not inner_text: the displayed text is cut off with an
    # ellipsis by the styling, but the DOM still holds all of it.
    texts = []
    try:
        texts.append(root.evaluate("el => el.textContent || ''"))
    except Exception:  # noqa: BLE001
        pass
    for frame in _frames(page):
        for getter in (
            lambda f=frame: f.evaluate("() => document.body.textContent || ''"),
            lambda f=frame: f.content(),
        ):
            try:
                texts.append(getter())
            except Exception:  # noqa: BLE001
                continue

    import html as _html

    for blob in texts:
        if not blob:
            continue
        blob = _html.unescape(blob)
        match = IFRAME_RE.search(blob)
        if match and "ispring" in match.group(0).lower() and match.group(0).endswith(">"):
            return match.group(0).strip()
        url = EMBED_URL_RE.search(blob)
        if url:
            _note("built the iframe from the embed URL")
            return EMBED_TEMPLATE.format(url=url.group(0))

    # The preview link and the embed player share an id.
    for selector in ("input[type=text]", "input:not([type])"):
        try:
            locator = root.locator(selector)
            count = locator.count()
        except Exception:  # noqa: BLE001
            count = 0
        for index in range(min(count, 12)):
            try:
                value = locator.nth(index).input_value(timeout=2000) or ""
            except Exception:  # noqa: BLE001
                continue
            preview = PREVIEW_URL_RE.search(value)
            if preview:
                url = f"{preview.group(1)}/app/embed-player/{preview.group(2)}"
                _note("built the iframe from the preview link", url=url)
                return EMBED_TEMPLATE.format(url=url)
    for blob in texts:
        preview = PREVIEW_URL_RE.search(_html.unescape(blob or ""))
        if preview:
            url = f"{preview.group(1)}/app/embed-player/{preview.group(2)}"
            _note("built the iframe from the preview link", url=url)
            return EMBED_TEMPLATE.format(url=url)
    return ""


def close_popup(page) -> None:
    """Close the share popup so the browser is clean for the next job.

    The X has no text, so it is found by its aria-label or as the only small
    button at the top of the popup. Escape alone is not always enough.
    """
    for attempt in range(3):
        if popup_root(page) is None:
            _note("share popup closed")
            return
        closer = visible_or_none(page.locator(POPUP_CLOSE))
        if closer is not None and safe_click(closer, "popup close (X)"):
            page.wait_for_timeout(900)
            if popup_root(page) is None:
                _note("share popup closed")
                return
        try:
            page.keyboard.press("Escape")
            page.wait_for_timeout(900)
        except Exception:  # noqa: BLE001
            pass
        if popup_root(page) is None:
            _note("share popup closed")
            return

        root = scope(page)
        for getter in (
            lambda: root.get_by_role("button", name=re.compile(r"close|dismiss", re.I)),
            lambda: root.locator("[aria-label*='lose' i]"),
            lambda: root.locator("button"),
        ):
            try:
                found = visible_or_none(getter())
            except Exception:  # noqa: BLE001
                found = None
            if found is not None and safe_click(found, "popup close"):
                page.wait_for_timeout(900)
                break
        if popup_root(page) is None:
            _note("share popup closed")
            return
        _warn("share popup still open", attempt=attempt + 1)

    # Last resort: reload the library, which drops any popup with it.
    try:
        page.reload(wait_until="domcontentloaded", timeout=30000)
        _note("reloaded the page to clear the popup")
    except Exception:  # noqa: BLE001
        _warn("share popup did not close")


CHROME_PATHS = (
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
)


def find_chrome(configured: str = "") -> str:
    if configured and Path(configured).is_file():
        return configured
    for candidate in CHROME_PATHS:
        if Path(candidate).is_file():
            return candidate
    import shutil as _shutil

    return _shutil.which("chrome") or _shutil.which("msedge") or ""


def debug_port_open(cdp_url: str, timeout_s: float = 2) -> bool:
    import urllib.request

    try:
        with urllib.request.urlopen(f"{cdp_url.rstrip('/')}/json/version", timeout=timeout_s):
            return True
    except Exception:  # noqa: BLE001
        return False


def start_browser(cdp_url: str, profile_dir: str, chrome_path: str = "") -> bool:
    """Start the browser the share step attaches to.

    It has to run on its own profile directory: since Chrome 136 the remote
    debugging port is ignored on the default profile, so the signed-in everyday
    Chrome cannot be used no matter what. This profile keeps its own session,
    which is why signing in once is enough.
    """
    import subprocess
    from urllib.parse import urlsplit

    executable = find_chrome(chrome_path)
    if not executable:
        _warn("no Chrome found to start; set ISPRING_CHROME_PATH")
        return False
    port = urlsplit(cdp_url).port or 9222
    Path(profile_dir).mkdir(parents=True, exist_ok=True)
    _note("starting the share browser", executable=executable, port=port)
    try:
        subprocess.Popen(
            [
                executable,
                f"--remote-debugging-port={port}",
                f"--user-data-dir={profile_dir}",
                "--no-first-run",
                "--no-default-browser-check",
            ],
            close_fds=True,
        )
    except Exception as exc:  # noqa: BLE001
        _warn("could not start the share browser", error=str(exc))
        return False
    for _ in range(30):
        if debug_port_open(cdp_url):
            _note("share browser is ready")
            return True
        time.sleep(1)
    _warn("the share browser did not open its debug port in time")
    return False


SEARCH_BOX_SELECTORS = (
    SEARCH_INPUT,  # the library's own box, read off the live page
    "input[type=search]",
    "input[placeholder*='Search' i]",
    "input[aria-label*='Search' i]",
    "[role=searchbox]",
    "input[name*='search' i]",
)


def search_library(page, term: str) -> bool:
    """Type the term into the library's search box and press Enter.

    The box is ``input[data-at="id=search-global-input"]`` at the top of the
    page. Typing is the way the site is meant to be used, so it is what this
    does; asking for the results URL is only a fallback for when the box is
    not on screen at all.

    Returns True when the results have been asked for — not that the material
    is in them. The caller checks that.
    """
    # A popup over the library swallows the click. Close it and carry on:
    # giving up here is what made the search never run at all.
    if popup_root(page) is not None:
        _warn("a popup is covering the library; closing it before searching")
        close_popup(page)

    tried: list[str] = []
    for frame in _frames(page):
        for selector in SEARCH_BOX_SELECTORS:
            try:
                box = visible_or_none(frame.locator(selector))
            except Exception as exc:  # noqa: BLE001
                tried.append(f"{selector}: {exc.__class__.__name__}")
                continue
            if box is None:
                tried.append(f"{selector}: none visible")
                continue
            try:
                box.click(timeout=4000)
                page.wait_for_timeout(300)
                box.fill("")
                page.wait_for_timeout(300)
                box.type(term, delay=60)
                page.wait_for_timeout(600)
                box.press("Enter")
                page.wait_for_timeout(3000)
                wait_for_rows(page)
                typed = ""
                try:
                    typed = box.input_value(timeout=2000) or ""
                except Exception:  # noqa: BLE001
                    pass
                _note(
                    "searched the library",
                    term=term,
                    selector=selector,
                    box_holds=typed,
                    results=len(row_titles(page, limit=5)),
                )
                return True
            except Exception as exc:  # noqa: BLE001
                tried.append(f"{selector}: {exc.__class__.__name__}")
                continue

    # No box to type into. The site keeps the query in the URL as base64 of
    # the escaped term, so this asks for the same results page directly.
    _warn("no search box on the page; asking for the results URL",
          term=term, attempts=tried[:8])
    encoded = base64.b64encode(quote(term).encode("utf-8")).decode("ascii")
    target = f"{site_root(page)}/app/s?s=search%2F{encoded}"
    try:
        page.goto(target, wait_until="domcontentloaded", timeout=45000)
        page.wait_for_timeout(2500)
        wait_for_rows(page)
        _note("searched by URL", term=term, url=target)
        return True
    except Exception as exc:  # noqa: BLE001
        _warn("search by URL failed too", term=term, error=str(exc))
    return False


def open_result(page, name: str) -> bool:
    """Click a row in the search results to open it.

    Searching for the institution returns its folder, not the material inside
    it, so the folder has to be opened before the material can be found.
    """
    pattern = re.compile(re.escape(name), re.I)
    for frame in _frames(page):
        row = visible_or_none(frame.get_by_text(pattern))
        if row is None:
            continue
        for action in ("click", "dblclick"):
            try:
                getattr(row, action)(timeout=4000)
                page.wait_for_timeout(2500)
                _note("opened search result", name=name, how=action)
                return True
            except Exception:  # noqa: BLE001
                continue
        if safe_click(row, f"search result {name}"):
            page.wait_for_timeout(2500)
            return True
    return False


RECENT_RE = re.compile(r"^\s*recent\b", re.I)
RECENT_PATH = "/app/s?s=recent"


def site_root(page, cloud_url: str = "") -> str:
    """The library's base URL, from the page we are on or the setting."""
    raw = cloud_url or page.url or ""
    return raw.split("/app/")[0].rstrip("/")


def go_recent(page, cloud_url: str = "") -> bool:
    """Load Recent from scratch.

    Loading the URL rather than clicking the menu matters: the tab has often
    been open since before the deck was published, so whatever it shows is
    out of date and the new material simply is not in it. A fresh load is
    the refresh.
    """
    target = f"{site_root(page, cloud_url)}{RECENT_PATH}"
    try:
        page.goto(target, wait_until="domcontentloaded", timeout=45000)
    except Exception as exc:  # noqa: BLE001
        _warn("could not load Recent; reloading this page instead",
              url=target, error=str(exc))
        try:
            page.reload(wait_until="domcontentloaded", timeout=45000)
            page.wait_for_timeout(3000)
        except Exception:  # noqa: BLE001
            pass
        return False
    page.wait_for_timeout(3500)
    _note("loaded Recent", url=target)
    return True


def open_recent(page) -> bool:
    """Open the library's Recent view, where the newest material is first.

    The deck was published seconds ago, so it is at the top of that list —
    no searching and no scrolling.
    """
    if popup_root(page) is not None:
        return False
    if not click_by_role(
        page, RECENT_RE, roles=("link", "button", "menuitem", "tab", "treeitem"),
        what="Recent",
    ):
        return False
    page.wait_for_timeout(2500)
    _note("opened Recent")
    return True


def go_library(page, cloud_url: str = "") -> bool:
    """Load the library from scratch, so nothing on screen is stale."""
    target = cloud_url or f"{site_root(page)}/"
    try:
        page.goto(target, wait_until="domcontentloaded", timeout=45000)
        page.wait_for_timeout(3500)
        _note("loaded the library", url=target)
        return True
    except Exception as exc:  # noqa: BLE001
        _warn("could not load the library; reloading", url=target, error=str(exc))
        try:
            page.reload(wait_until="domcontentloaded", timeout=45000)
            page.wait_for_timeout(3000)
        except Exception:  # noqa: BLE001
            return False
        return True


def locate_material(
    page, material: str, institution: str, cloud_url: str = "", rounds: int = 3
) -> bool:
    """Find the just-published material's row.

    The way that works, in this order:

    1. search for the **institution**, open the folder the search returns,
       and scroll down it until the material appears;
    2. search for the material by name;
    3. Recent, where the newest sits at the top;
    4. walk the library by hand.

    Every round starts by loading the page again. A tab that has been open
    since before the publish shows an old list, and iSpring Cloud takes a
    little while to list a material anyway — so a round that finds nothing
    waits and looks again rather than failing.
    """
    for round_no in range(1, max(1, rounds) + 1):
        # 1. Search for the institution and open the folder it returns, then
        # look inside it. This is the route to prefer when the project exists.
        if institution and search_library(page, institution):
            if open_result(page, institution) and find_row(page, material, institution=institution) is not None:
                _note("found it in the institution folder",
                      folder=institution, round=round_no)
                return True

        # 2. Search for the material itself.
        if search_library(page, material) and find_row(page, material, institution=institution) is not None:
            _note("found it by name", material=material, round=round_no)
            return True

        # 3. Recent: the newest is at the top. A deck that went to the parent
        # folder because the institution has no project of its own is here.
        if go_recent(page, cloud_url) and find_row(page, material, institution=institution) is not None:
            _note("found it in Recent", material=material, round=round_no)
            return True

        # 4. No search at all: open the folder from the list and scroll.
        if institution:
            go_library(page, cloud_url)
            if enter_folder(page, institution) and find_row(page, material, institution=institution) is not None:
                _note("found it by walking the library", folder=institution)
                return True

        _warn(
            "not listed yet",
            material=material,
            round=round_no,
            on_screen=row_titles(page, limit=15) or visible_row_labels(page, limit=10),
        )
        if round_no < rounds:
            page.wait_for_timeout(10000)
    return False


COVER_RE = re.compile(r"edit cover image", re.I)
COVER_DIALOG_RE = re.compile(r"cover image settings", re.I)
SAVE_RE = re.compile(r"^\s*save\s*$", re.I)
CANCEL_RE = re.compile(r"^\s*cancel\s*$", re.I)
TITLE_LABEL_RE = re.compile(r"^\s*title\s*$", re.I)


def visible_menu_labels(page) -> list[str]:
    """What is on screen right now, for when an expected item is missing."""
    labels = []
    for frame in _frames(page):
        for role in ("menuitem", "option", "button", "link"):
            try:
                locator = frame.get_by_role(role)
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
                if label and label not in labels:
                    labels.append(label)
    return labels


def popup_with_text(page, pattern):
    """The popup that contains this text, when several are stacked."""
    for frame in _frames(page):
        for selector in POPUP_SELECTORS:
            try:
                locator = frame.locator(selector)
                count = locator.count()
            except Exception:  # noqa: BLE001
                continue
            for index in range(min(count, 6)):
                item = locator.nth(index)
                try:
                    if not item.is_visible():
                        continue
                    if visible_or_none(item.get_by_text(pattern)) is not None:
                        return item
                except Exception:  # noqa: BLE001
                    continue
    return None


def open_share(page, material: str) -> bool:
    """Open the row's three-dot menu and click Share.

    Uses the library's own markup: find the row by its title, hover it so the
    three dots appear, open the menu and click Share inside that menu — not
    anywhere on the page, which is how a click used to land on the row behind.
    """
    row = find_row(page, material)
    if row is None:
        _warn("no row for the material", material=material, on_screen=row_titles(page))
        return False

    for attempt in range(1, 4):
        if not open_row_menu_for(page, row):
            page.wait_for_timeout(1000)
            continue
        menu = page.locator(ROW_MENU_POPOVER).first
        share = visible_or_none(menu.locator(SHARE_MENU_ITEM))
        if share is None:
            share = visible_or_none(menu.get_by_text(SHARE_RE))
        if share is not None and safe_click(share, "Share"):
            page.wait_for_timeout(1500)
            return True
        items = []
        try:
            items = [
                line.strip()
                for line in (menu.inner_text() or "").splitlines()
                if line.strip()
            ]
        except Exception:  # noqa: BLE001
            pass
        _warn("no Share in the row menu", attempt=attempt, items=items[:15])
        try:
            page.keyboard.press("Escape")
            page.wait_for_timeout(600)
        except Exception:  # noqa: BLE001
            pass

    # Last resort: the old way, in case the markup has changed under us.
    if open_row_menu(page, material):
        page.wait_for_timeout(1200)
        if click_by_role(page, SHARE_RE, what="Share"):
            return True
    return False


def set_cover_title(page, title: str) -> bool:
    """Rename the cover image title to the material name.

    iSpring names the cover after the file it published, which is the working
    copy — "source" or "source.repaired" — so it has to be corrected here.
    """
    cover = visible_or_none(page.locator(COVER_BUTTON))
    if cover is not None:
        safe_click(cover, "Edit cover image")
    elif not click_by_role(page, COVER_RE, roles=("button", "link"), what="Edit cover image"):
        _warn("no Edit cover image button; leaving the cover title alone")
        return False
    page.wait_for_timeout(2000)

    dialog = popup_with_text(page, COVER_DIALOG_RE)
    if dialog is None:
        _warn("the cover image dialog did not open")
        return False

    box = None
    for selector in ("input[type=text]", "input:not([type])", "textarea"):
        try:
            box = visible_or_none(dialog.locator(selector))
        except Exception:  # noqa: BLE001
            box = None
        if box is not None:
            break
    if box is None:
        _warn("no title box in the cover image dialog")
        return False

    # Already right: close the dialog without saving. Saving an unchanged
    # title is a pointless write, and it makes iSpring rebuild the cover.
    try:
        current = (box.input_value(timeout=3000) or "").strip()
    except Exception:  # noqa: BLE001
        current = ""
    if current and current.casefold() == title.strip().casefold():
        _note("cover title is already the material name; cancelling", title=current)
        for locator in (
            dialog.get_by_role("button", name=CANCEL_RE),
            dialog.get_by_text(CANCEL_RE),
        ):
            found = visible_or_none(locator)
            if found is not None and safe_click(found, "Cancel (cover title)"):
                page.wait_for_timeout(1200)
                return True
        try:
            page.keyboard.press("Escape")
            page.wait_for_timeout(1200)
        except Exception:  # noqa: BLE001
            pass
        return True

    try:
        box.click(timeout=4000)
        box.fill("")
        box.type(title, delay=25)
        page.wait_for_timeout(800)
    except Exception as exc:  # noqa: BLE001
        _warn("could not type the cover title", error=str(exc))
        return False

    saved = False
    for locator in (dialog.get_by_role("button", name=SAVE_RE), dialog.get_by_text(SAVE_RE)):
        found = visible_or_none(locator)
        if found is not None and safe_click(found, "Save (cover title)"):
            saved = True
            break
    if not saved:
        _warn("could not press Save in the cover image dialog")
        return False

    page.wait_for_timeout(2000)
    _note("set the cover title", title=title)
    return True


# --- creating a folder for an institution that has none ---------------------
#
# Read from the site's own APIs, write through its UI. The create request's
# body is not documented anywhere we can see, and a private API guessed at is
# the kind of thing that fails quietly six months from now; the two GETs below
# are only lookups, so a change there fails loudly instead.

PROJECT_LIST_API = "/s/api/v1/project/list"
FOLDER_LIST_API = "/s/api/v1/content/list/project?folderID="
ADD_BUTTON = '[data-at="id=add-material-button"]'
ADD_FOLDER_ITEM = '[data-at="id=add-material-folder"]'
CREATE_FOLDER_POPUP = '[data-at="id=create-folder-popup"]'
FOLDER_NAME_INPUT = '[data-at="id=folder-name-input"]'
CREATE_FOLDER_CONFIRM = '[data-at="id=create-folder-button"]'


def _cloud_json(page, path: str):
    """Call one of the library's own endpoints from the signed-in page.

    A signed-out browser answers 401 here rather than anything useful, so the
    failure is named: that is the one a person can actually fix.
    """
    try:
        return _cloud_fetch(page, path)
    except Exception as exc:  # noqa: BLE001
        raise ISpringPublishingError(
            f"could not read {path} from iSpring Cloud: {exc}",
            user_message=(
                "Could not read the iSpring Cloud library. The browser it uses "
                "may have been signed out; sign in once in that Chrome window."
            ),
        ) from exc


def _cloud_fetch(page, path: str):
    return page.evaluate(
        """async (path) => {
            const response = await fetch(path, {credentials: 'include'});
            if (!response.ok) throw new Error('HTTP ' + response.status);
            return await response.json();
        }""",
        path,
    )


def folder_titles(page, root_folder: str) -> list:
    listing = _cloud_json(page, FOLDER_LIST_API + root_folder)
    return [
        normalise(item.get("title", ""))
        for item in (listing.get("content") or [])
        if item.get("type") == "FOLDER"
    ]


def ensure_project_folder(
    institution: str,
    parent_folder: str,
    cdp_url: str,
    profile_dir: str = r"C:\ispring-chrome-profile",
    chrome_path: str = "",
    cloud_url: str = "https://harshit.ispring.com/",
) -> bool:
    """Create the institution's folder under ``parent_folder``, if it is missing.

    Returns True when a folder was created, False when one was already there.

    Done in the browser because the Suite publish dialog has no way to make a
    folder. The name box opens filled in with "New Folder", so it is cleared
    and typed into, and the result is read back from the library afterwards:
    a folder created under the wrong name is worse than none at all, and that
    is exactly what a half-registered keystroke produces.
    """
    from playwright.sync_api import sync_playwright

    wanted = normalise(institution)
    parent = normalise(parent_folder)
    if not wanted or not parent:
        raise ISpringPublishingError(
            "ensure_project_folder needs both an institution and a parent folder",
            user_message=USER_PUBLISH_FAILED,
        )

    if not debug_port_open(cdp_url):
        _note("no browser on the debug port; starting one")
        start_browser(cdp_url, profile_dir, chrome_path)

    with sync_playwright() as playwright:
        try:
            browser = playwright.chromium.connect_over_cdp(cdp_url)
        except Exception as exc:  # noqa: BLE001
            raise ISpringPublishingError(
                f"could not attach to Chrome at {cdp_url}: {exc}",
                user_message=(
                    "The browser this step uses is not running. Start it with "
                    "scripts\\start_ispring_chrome.cmd, sign in to iSpring "
                    "Cloud once, and leave it open."
                ),
            ) from exc
        if not browser.contexts:
            raise ISpringPublishingError(
                "attached to Chrome but it has no windows open",
                user_message=USER_PUBLISH_FAILED,
            )
        context = browser.contexts[0]
        page = context.pages[-1] if context.pages else context.new_page()
        page.bring_to_front()

        try:
            page.goto(cloud_url, wait_until="domcontentloaded", timeout=60000)
        except Exception as exc:  # noqa: BLE001
            raise ISpringPublishingError(
                f"could not open {cloud_url}: {exc}",
                user_message=USER_PUBLISH_FAILED,
            ) from exc
        page.wait_for_timeout(4000)

        # {"projects": [...], "totalProjectsCount": n} — not a bare list.
        listing = _cloud_json(page, PROJECT_LIST_API)
        projects = listing.get("projects") if isinstance(listing, dict) else listing
        target = None
        for project in projects or []:
            if normalise(project.get("title", "")).lower() == parent.lower():
                target = project
                break
        if target is None:
            raise ISpringPublishingError(
                f"no project named {parent!r} in iSpring Cloud",
                user_message=(
                    f"The parent folder {parent!r} was not found in iSpring "
                    "Cloud. Check ISPRING_NEW_INSTITUTION_PARENT."
                ),
            )

        root = str(target.get("rootFolder") or "")
        existing = folder_titles(page, root)
        if any(title.lower() == wanted.lower() for title in existing):
            _note("the folder already exists", institution=wanted, parent=parent)
            return False

        # "Esromagica" next to "Esro Magica", "LawSikho" next to "Law Sikho":
        # seven pairs like that are already in the library. Creating an eighth
        # splits one institution's decks across two folders and nobody notices
        # for months, so this stops and says which folder it means.
        squashed = wanted.lower().replace(" ", "")
        near = [
            title for title in existing
            if title.lower().replace(" ", "") == squashed
        ]
        if near:
            raise ISpringPublishingError(
                f"{parent!r} already holds {near[0]!r}, which is {wanted!r} "
                "without the spaces",
                user_message=(
                    f"A folder called {near[0]!r} already exists in {parent!r}. "
                    f"Rename it to {wanted!r}, or fix the institution name, "
                    "rather than having both."
                ),
            )

        # Get into the project, and prove we are in it before touching Add.
        #
        # The library is a single-page app: goto loads the shell and the app
        # then routes itself, so for a second or two the screen still shows
        # whatever it showed before — usually the Learning Content root. The
        # Add button there belongs to that view, and clicking it creates the
        # folder at the top of the library instead of inside the parent.
        # Nothing in the page says which view you are on; the address does.
        project_id = str(target.get("id") or "")
        target_url = f"{cloud_url.rstrip('/')}/app/s?s=project%2F{project_id}%2F{root}"
        landed = False
        for attempt in range(1, 4):
            try:
                page.goto(target_url, wait_until="domcontentloaded", timeout=60000)
            except Exception as exc:  # noqa: BLE001
                _warn("could not open the project page", attempt=attempt, error=str(exc))
                continue
            for _ in range(15):
                page.wait_for_timeout(1000)
                here = (page.evaluate("() => location.href") or "").lower()
                if project_id.lower() in here and root.lower() in here:
                    landed = True
                    break
            if landed:
                break
            _warn("the library did not stay on the project page", attempt=attempt)

        if not landed:
            raise ISpringPublishingError(
                f"could not open {parent!r} ({target_url}); refusing to create a "
                "folder while the library is showing something else",
                user_message=(
                    f"Could not open the {parent!r} folder in iSpring Cloud, so "
                    f"no folder was created for {wanted!r}."
                ),
            )

        # The Add button has to belong to the project view, not a half-drawn
        # previous one.
        page.wait_for_selector(ADD_BUTTON, timeout=30000)
        page.wait_for_timeout(1500)

        page.click(ADD_BUTTON, timeout=20000)
        page.click(ADD_FOLDER_ITEM, timeout=20000)
        page.wait_for_selector(CREATE_FOLDER_POPUP, timeout=20000)

        box = page.locator(FOLDER_NAME_INPUT)
        box.click(timeout=10000)
        box.fill("")
        box.type(wanted, delay=25)
        page.wait_for_timeout(500)
        if normalise(box.input_value()) != wanted:
            raise ISpringPublishingError(
                f"the name box holds {box.input_value()!r}, not {wanted!r}",
                user_message=USER_PUBLISH_FAILED,
            )

        page.click(CREATE_FOLDER_CONFIRM, timeout=20000)
        page.wait_for_timeout(4000)

        after = folder_titles(page, root)
        if not any(title.lower() == wanted.lower() for title in after):
            raise ISpringPublishingError(
                f"a folder was created but {wanted!r} is not in {parent!r} "
                "afterwards — check the top of the library, it may be sitting "
                "there and will need deleting",
                user_message=(
                    f"Could not create the iSpring Cloud folder for {wanted!r} "
                    f"inside {parent!r}. Check the library for a stray folder."
                ),
            )
        _note("created the institution folder", institution=wanted, parent=parent)
        return True


def fill_container(embed: str) -> str:
    """Make the iframe fill whatever holds it, instead of 1280x720.

    Only the opening tag is touched, and only its width and height: every
    other attribute iSpring puts there — allowfullscreen, the frameborder,
    the src — is left exactly as it came.
    """

    def rewrite(match: "re.Match[str]") -> str:
        tag = SIZE_ATTR_RE.sub("", match.group(0))
        return tag.replace("<iframe", '<iframe width="100%" height="100%"', 1)

    return IFRAME_OPEN_RE.sub(rewrite, embed, count=1)


def fetch_embed(
    material: str,
    institution: str,
    cdp_url: str,
    profile_dir: str = r"C:\ispring-chrome-profile",
    chrome_path: str = "",
    cloud_url: str = "https://harshit.ispring.com/",
    cover_title: str = "",
) -> str:
    """Drive the Chrome that is already running and already signed in.

    Nothing is launched and no profile is created, so a login done once by
    hand in that Chrome keeps working for every run after it.
    """
    from playwright.sync_api import sync_playwright

    if not debug_port_open(cdp_url):
        _note("no browser on the debug port; starting one")
        start_browser(cdp_url, profile_dir, chrome_path)

    with sync_playwright() as playwright:
        try:
            browser = playwright.chromium.connect_over_cdp(cdp_url)
        except Exception as exc:  # noqa: BLE001
            raise ISpringPublishingError(
                f"could not attach to Chrome at {cdp_url}: {exc}",
                user_message=(
                    "The browser the share step uses is not running. Start it "
                    "with scripts\\start_ispring_chrome.cmd, sign in to iSpring "
                    "Cloud once, and leave it open."
                ),
            ) from exc
        if not browser.contexts:
            raise ISpringPublishingError(
                "attached to Chrome but it has no windows open",
                user_message=USER_LINK_FAILED,
            )
        context = browser.contexts[0]

        # Manage Content opens the machine's default browser, which is not
        # necessarily this one, so an iSpring tab may never appear here. Give
        # it a moment, then open the library directly rather than depend on it.
        page = None
        for _ in range(8):
            pages = [p for p in context.pages if "ispring" in (p.url or "").lower()]
            if pages:
                page = pages[-1]
                break
            time.sleep(1)
        if page is None:
            _note("no iSpring tab yet; opening the library", url=cloud_url)
            page = context.new_page() if not context.pages else context.pages[-1]
            try:
                page.goto(cloud_url, wait_until="domcontentloaded", timeout=60000)
            except Exception as exc:  # noqa: BLE001
                raise ISpringPublishingError(
                    f"could not open {cloud_url}: {exc}",
                    user_message=USER_LINK_FAILED,
                ) from exc
        page.bring_to_front()
        page.wait_for_timeout(3000)

        if visible_or_none(page.get_by_text(re.compile(r"sign in|log in", re.I))) is not None:
            _warn("that Chrome may not be signed in to iSpring Cloud")

        _note("browser is on", url=(page.url or "")[:200])
        if dump_always():
            dump_page(page, "library")

        if not locate_material(page, material, institution, cloud_url=cloud_url):
            dump = dump_page(page, "material-not-found")
            raise ISpringPublishingError(
                f"{material!r} was not found in the library (url {page.url}); "
                f"browser dump: {dump or 'none'}",
                user_message=(
                    f"The published material {material!r} could not be found in "
                    "iSpring Cloud to read its share link."
                ),
            )

        if not open_share(page, material):
            dump = dump_page(page, "no-share-menu")
            raise ISpringPublishingError(
                f"no Share item in the row menu; on screen: "
                f"{visible_menu_labels(page)[:20]}; browser dump: {dump or 'none'}",
                user_message=USER_LINK_FAILED,
            )
        page.wait_for_timeout(2000)

        if not make_viewable_via_link(page):
            dump = dump_page(page, "toggle-not-switched-on")
            raise ISpringPublishingError(
                "'Make viewable via link' could not be switched on; the embed "
                f"code would point at content nobody can open; browser dump: "
                f"{dump or 'none'}",
                user_message=USER_LINK_FAILED,
            )
        page.wait_for_timeout(2000)

        # The deck is published under its id, so the cover would read "88214"
        # unless it is set to the name a person expects to see.
        set_cover_title(page, cover_title or material)

        embed = read_embed_code(page)
        if not embed:
            click_by_role(page, EMBED_RE, roles=("tab", "button", "link"), what="Embed")
            page.wait_for_timeout(1500)
            embed = read_embed_code(page)

        dump = "" if embed else dump_page(page, "no-embed-code")
        if dump_always():
            dump_page(page, "share-popup")
        close_popup(page)
        if not embed:
            raise ISpringPublishingError(
                f"the share popup showed no iframe code; browser dump: {dump or 'none'}",
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
    parent_folders: Sequence[str] = ("PPT Migration", "PPT migration New"),
    cdp_url: str = "http://127.0.0.1:9222",
    publish_timeout_s: float = 1800,
    close_powerpoint_after: bool = True,
    skip_open: bool = False,
    browser_profile_dir: str = r"C:\ispring-chrome-profile",
    browser_path: str = "",
    cloud_url: str = "https://harshit.ispring.com/",
    press_manage_content: bool = False,
    cover_title: str = "",
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
        window = find_powerpoint_window(pptx)
        # The picker is clicked with a real mouse click, which lands wherever
        # the front window is. The folder-creation step puts Chrome in front,
        # so on the retry PowerPoint has to be brought back first.
        try:
            window.set_focus()
            time.sleep(1)
        except Exception:  # noqa: BLE001
            _warn("could not bring PowerPoint to the front")
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

        pick_project(dialog, institution, parent_folders)

        publish = dialog.child_window(auto_id=ID_OK, control_type="Button")
        if not publish.exists():
            raise ISpringPublishingError(
                "Publish button not found on the publish dialog",
                user_message=USER_PUBLISH_FAILED,
            )
        activate(publish, "Publish")
        wait_for_completion(window, publish_timeout_s)

        # Manage Content is deliberately not pressed: it opens Windows' default
        # browser, which is not the one the share step drives. It would steal
        # focus from the automation and leave a stray window behind every job.
        if press_manage_content:
            click_manage_content(window)
            time.sleep(10)
        else:
            close_ispring_windows(window)

        # "Publishing is complete!" means PowerPoint has finished; iSpring is
        # still uploading the content to the cloud for a few seconds after
        # that. Going to the browser too early finds a library without it.
        _note("waiting for the upload to finish", seconds=UPLOAD_SETTLE_S)
        time.sleep(UPLOAD_SETTLE_S)

        embed = fetch_embed(
            content_name or institution,
            institution,
            cdp_url,
            profile_dir=browser_profile_dir,
            chrome_path=browser_path,
            cloud_url=cloud_url,
            cover_title=cover_title,
        )
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

    embed = fill_container(embed)
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
