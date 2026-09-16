"""Read-only iSpring Cloud web probe.

This module drives the iSpring Cloud website in a browser. It does not
automate PowerPoint, iSpring Suite, or the worker poll loop.
"""

from __future__ import annotations

import json
import logging
import re
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Mapping, Sequence
from urllib.parse import parse_qs, unquote, urlencode, urlsplit

logger = logging.getLogger("ispring_cloud")

OFFICIAL_CLOUD_APP_URL = "https://pro.ispringcloud.com/"
OFFICIAL_CLOUD_MARKETING_URL = "https://www.ispringsolutions.com/ispring-cloud"
OFFICIAL_ID_LOGIN_URL = (
    "https://id.ispring.com/login?service=isa&redirect_url="
    "https%3A%2F%2Fwww.ispringsolutions.com%2Faccount%2Flogin-by-isid&lang=en"
)
DEFAULT_TEAM_APP_URL = (
    "https://harshit.ispring.com/app/s?s=project%2F"
    "65dcde53-2c22-11ec-b427-a2fdffd5f86a%2F"
    "45ed05bc-2c20-11ec-b54a-76ca0a89c886"
)
UUID_RE = re.compile(
    r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}",
    re.IGNORECASE,
)
OFFICIAL_HOST_SUFFIXES = (
    "ispringcloud.com",
    "ispringsolutions.com",
    "ispring.com",
)
DEFAULT_SEARCH_QUERY = "new ppt migration"
DEFAULT_MATERIAL_NAME = "demo academy"
SITE_SEARCH_QUERY = "iSpring Cloud official"

NAV_TIMEOUT_MS = 60_000
ACTION_TIMEOUT_MS = 15_000
SEARCH_TIMEOUT_MS = 30_000

COOKIE_BUTTON_NAMES = (
    "Accept All",
    "Accept all",
    "Accept All Cookies",
    "Accept cookies",
    "Accept",
    "I agree",
    "Agree",
    "Allow all",
    "Allow cookies",
    "Got it",
    "OK",
    "Close",
)

MATERIAL_QUERY_KEYS = frozenset(
    {
        "contentid",
        "content_id",
        "materialid",
        "material_id",
        "itemid",
        "item_id",
        "presentationid",
        "presentation_id",
        "id",
    }
)
PATH_ID_MARKERS = frozenset(
    {
        "content",
        "contents",
        "item",
        "items",
        "presentation",
        "presentations",
        "material",
        "materials",
        "s",
        "p",
        "acc",
    }
)
GENERIC_PATH_SEGMENTS = frozenset(
    {
        "",
        "app",
        "login",
        "signin",
        "sign-in",
        "auth",
        "account",
        "projects",
        "project",
        "content",
        "contents",
        "home",
        "library",
        "search",
        "www",
        "pro",
        "index.html",
    }
)
DESTRUCTIVE_TOKENS = frozenset(
    {
        "delete",
        "upload",
        "publish",
        "unpublish",
        "edit",
        "remove",
        "rename",
        "archive",
        "overwrite",
        "replace",
        "trash",
        "create",
    }
)
LOGIN_PATH_RE = re.compile(
    r"/(login|log-in|signin|sign-in|auth)(/|$)", re.IGNORECASE
)

PauseFn = Callable[[str], object]


class ISpringCloudProbeError(RuntimeError):
    """Raised when the read-only cloud probe cannot continue."""


@dataclass(frozen=True)
class MaterialInfo:
    title: str
    url: str
    material_id: str | None
    share_controls: tuple[str, ...]


@dataclass
class ArtifactRun:
    root: Path
    log_path: Path

    @classmethod
    def create(cls, base: Path) -> ArtifactRun:
        stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        root = Path(base) / "ispring-cloud" / stamp
        root.mkdir(parents=True, exist_ok=True)
        log_path = root / "session.log"
        log_path.touch()
        return cls(root=root, log_path=log_path)

    def screenshot_name(self, step: str) -> str:
        safe = re.sub(r"[^a-z0-9]+", "-", step.casefold()).strip("-")
        return f"{safe}.png"

    def write_log(self, message: str) -> Path:
        stamp = datetime.now(timezone.utc).isoformat()
        with self.log_path.open("a", encoding="utf-8") as handle:
            handle.write(f"{stamp} {message}\n")
        return self.log_path

    def write_text(self, name: str, content: str) -> Path:
        path = self.root / name
        path.write_text(content, encoding="utf-8")
        return path


def normalize_text(value: str) -> str:
    return " ".join(value.split()).casefold()


def text_matches(candidate: str, target: str) -> bool:
    if not candidate or not target:
        return False
    candidate_n = normalize_text(candidate)
    target_n = normalize_text(target)
    if target_n in candidate_n:
        return True
    return target_n.replace(" ", "") in candidate_n.replace(" ", "")


def name_pattern(name: str) -> re.Pattern[str]:
    parts = [re.escape(part) for part in name.split() if part]
    if not parts:
        raise ValueError("name must contain visible text")
    return re.compile(r"\s*".join(parts), re.IGNORECASE)


def is_team_app_url(url: str) -> bool:
    host = (urlsplit(url).hostname or "").casefold()
    if not host:
        return False
    if host.endswith("ispringcloud.com"):
        return True
    if host.endswith(".ispring.com") and host not in {
        "id.ispring.com",
        "www.ispring.com",
    }:
        return True
    return False


def is_official_ispring_url(url: str) -> bool:
    try:
        host = (urlsplit(url).hostname or "").casefold()
    except ValueError:
        return False
    if not host:
        return False
    return any(
        host == suffix or host.endswith("." + suffix)
        for suffix in OFFICIAL_HOST_SUFFIXES
    )


def _result_rank(url: str) -> int:
    host = (urlsplit(url).hostname or "").casefold()
    path = urlsplit(url).path.casefold()
    if host == "pro.ispringcloud.com":
        return 0
    if host.endswith("ispringcloud.com"):
        return 1
    if host.endswith("ispringsolutions.com") and "ispring-cloud" in path:
        return 2
    if host.endswith("ispringsolutions.com"):
        return 3
    if host.endswith("ispring.com"):
        return 4
    return 99


def pick_official_search_result(results: Sequence[tuple[str, str]]) -> str | None:
    official = [
        (title, url) for title, url in results if url and is_official_ispring_url(url)
    ]
    if not official:
        return None
    official.sort(key=lambda item: _result_rank(item[1]))
    return official[0][1]


def search_engine_url(query: str) -> str:
    return "https://duckduckgo.com/?" + urlencode({"q": query})


def _id_from_query(query: str) -> str | None:
    if not query:
        return None
    parsed = parse_qs(query, keep_blank_values=False)
    for key, values in parsed.items():
        if key.casefold() in MATERIAL_QUERY_KEYS and values and values[0].strip():
            return values[0].strip()
    for value in parsed.get("s") or []:
        decoded = unquote(value)
        uuids = UUID_RE.findall(decoded)
        if uuids:
            return uuids[-1]
        found = _id_from_path(decoded)
        if found:
            return found
    return None


def _id_from_path(path: str) -> str | None:
    segments = [segment for segment in path.split("/") if segment]
    if not segments:
        return None
    for index, segment in enumerate(segments[:-1]):
        if segment.casefold() not in PATH_ID_MARKERS:
            continue
        candidate = segments[index + 1].split("?")[0].strip()
        if candidate and candidate.casefold() not in GENERIC_PATH_SEGMENTS:
            return candidate
    return None


def extract_material_id(url: str) -> str | None:
    try:
        parts = urlsplit(url)
    except ValueError:
        return None
    found = _id_from_query(parts.query)
    if found:
        return found
    fragment = parts.fragment or ""
    if fragment:
        frag_path, _, frag_query = fragment.partition("?")
        found = _id_from_query(frag_query)
        if found:
            return found
        if "=" in frag_path and not frag_path.startswith("/"):
            found = _id_from_query(frag_path.lstrip("?#"))
            if found:
                return found
        found = _id_from_path(frag_path)
        if found:
            return found
    return _id_from_path(parts.path)


def looks_like_login(url: str, visible_text: str) -> bool:
    try:
        parts = urlsplit(url)
    except ValueError:
        parts = urlsplit("")
    if LOGIN_PATH_RE.search(parts.path or "") or LOGIN_PATH_RE.search(
        "/" + (parts.fragment or "").lstrip("#/")
    ):
        return True
    text = visible_text or ""
    signals = 0
    if re.search(r"\b(forgot password|welcome back|remember me)\b", text, re.I):
        signals += 1
    if re.search(r"\b(sign in|log in|signin|login)\b", text, re.I):
        signals += 1
    if re.search(r"\bpassword\b", text, re.I):
        signals += 1
    return signals >= 2


def is_destructive_label(label: str) -> bool:
    normalized = normalize_text(label)
    if not normalized:
        return False
    return bool(set(normalized.split()) & DESTRUCTIVE_TOKENS)


def login_credentials(environ: Mapping[str, str] | None = None) -> tuple[str, str] | None:
    source = environ if environ is not None else {}
    email = (source.get("ISPRING_CLOUD_EMAIL") or "").strip()
    password = source.get("ISPRING_CLOUD_PASSWORD") or ""
    if email and password:
        return email, password
    return None


def _find_login_field(page: Any, *, kind: str) -> Any:
    label = re.compile(kind, re.I)
    kind_key = kind.casefold()
    if kind_key == "email":
        strategies: tuple[Callable[[Any], Any], ...] = (
            lambda root: root.get_by_placeholder(re.compile(r"login|email", re.I)),
            lambda root: root.get_by_label(re.compile(r"login|email", re.I)),
            lambda root: root.locator("input[name='login' i]"),
            lambda root: root.locator("input[name='email' i]"),
            lambda root: root.locator("#emailLoginField"),
            lambda root: root.locator("input[type='email']"),
            lambda root: root.get_by_role("textbox").first,
        )
    else:
        strategies = (
            lambda root: root.locator("#passwordLoginField"),
            lambda root: root.locator("input[name='password' i][type='password']"),
            lambda root: root.locator("input[type='password']"),
            lambda root: root.get_by_label(label),
        )
    for frame in _frames(page):
        for strategy in strategies:
            found = _first_visible(strategy(frame))
            if found is not None:
                return found
    raise ISpringCloudProbeError(f"Could not find the {kind} field on the login page.")


def _find_sign_in_control(page: Any) -> Any | None:
    strategies: tuple[Callable[[Any], Any], ...] = (
        lambda root: root.get_by_role("button", name=re.compile(r"^continue$", re.I)),
        lambda root: root.get_by_role("button", name=re.compile(r"^sign in$", re.I)),
        lambda root: root.get_by_role("button", name=re.compile(r"^log in$", re.I)),
        lambda root: root.locator("input[type='submit']"),
        lambda root: root.locator("#btnLogin"),
        lambda root: root.get_by_text(re.compile(r"^(sign in|log in)$", re.I)),
    )
    for frame in _frames(page):
        for strategy in strategies:
            found = _first_visible(strategy(frame))
            if found is not None:
                label = _accessible_name(found)
                if re.search(r"google|facebook|linkedin|apple", label, re.I):
                    continue
                return found
    return None


def _fill_password_field(page: Any, password: str) -> None:
    """Type into the real password input, including iSpring's overlay widget."""
    overlay = page.locator("input.password_form_field_placeholder")
    try:
        if overlay.count() and overlay.first.is_visible():
            overlay.first.click(timeout=ACTION_TIMEOUT_MS)
    except Exception:  # noqa: BLE001
        logger.info("password placeholder overlay was not clickable")
    real = None
    for factory in (
        lambda: page.locator("#passwordLoginField"),
        lambda: page.locator("input[type='password']"),
        lambda: page.get_by_label(re.compile(r"password", re.I)),
    ):
        candidate = factory()
        found = _first_visible(candidate)
        if found is not None:
            real = found
            break
    if real is None:
        raise ISpringCloudProbeError("Could not find the password field on the login page.")
    real.click(timeout=ACTION_TIMEOUT_MS)
    try:
        real.fill("", timeout=ACTION_TIMEOUT_MS)
    except Exception:  # noqa: BLE001
        pass
    try:
        real.press_sequentially(password, delay=25)
    except Exception:  # noqa: BLE001
        page.keyboard.type(password, delay=25)
    entered = ""
    try:
        entered = real.input_value(timeout=ACTION_TIMEOUT_MS)
    except Exception:  # noqa: BLE001
        entered = ""
    if not entered:
        raise ISpringCloudProbeError(
            "Password field did not accept input. The password was not stored."
        )
    logger.info("password field accepted %s characters", len(entered))


def submit_login_form(page: Any, email: str, password: str) -> None:
    logger.info("submitting iSpring Cloud login form for %s", email)
    dismiss_cookie_dialogs(page)
    try:
        page.get_by_placeholder(re.compile(r"login|email", re.I)).first.wait_for(
            timeout=NAV_TIMEOUT_MS
        )
    except Exception:  # noqa: BLE001
        logger.info("Login/email placeholder not found; scanning locators")
    dismiss_cookie_dialogs(page)
    email_box = _find_login_field(page, kind="Email")
    email_box.click(timeout=ACTION_TIMEOUT_MS)
    email_box.fill(email, timeout=ACTION_TIMEOUT_MS)
    if _first_visible(page.locator("input[type='password']")) is None:
        logger.info("password field not on this step; clicking Continue")
        continue_btn = _find_sign_in_control(page)
        if continue_btn is None:
            email_box.press("Enter")
        else:
            continue_btn.click(timeout=ACTION_TIMEOUT_MS)
        deadline = time.monotonic() + 20
        while time.monotonic() < deadline:
            if not _page_is_open(page):
                logger.info("login tab closed after email step")
                return
            if _first_visible(page.locator("input[type='password']")) is not None:
                break
            time.sleep(0.3)
        else:
            raise ISpringCloudProbeError(
                "Password step did not appear after email Continue."
            )
    _fill_password_field(page, password)
    sign_in = _find_sign_in_control(page)
    if sign_in is None:
        page.locator("input[type='password']").press("Enter")
    else:
        sign_in.click(timeout=ACTION_TIMEOUT_MS)
    deadline = time.monotonic() + 45
    while time.monotonic() < deadline:
        if not _page_is_open(page):
            logger.info("login tab closed after submit; treating as a redirect")
            return
        try:
            page.wait_for_load_state("domcontentloaded", timeout=5000)
        except Exception:  # noqa: BLE001
            if not _page_is_open(page):
                logger.info("login tab closed during redirect")
                return
        dismiss_cookie_dialogs(page)
        try:
            visible = _visible_text(page)
            current_url = page.url
        except Exception:  # noqa: BLE001
            logger.info("login page became unavailable after submit")
            return
        if re.search(
            r"invalid email|invalid password|incorrect password|wrong password|"
            r"could not log in|authentication failed",
            visible,
            re.I,
        ):
            raise ISpringCloudProbeError(
                "iSpring Cloud rejected the login. The password was not stored."
            )
        if not looks_like_login(current_url, visible):
            logger.info("login form is no longer visible")
            return
        time.sleep(0.5)
    raise ISpringCloudProbeError(
        "Login form is still visible after submitting credentials. "
        "Check email/password, SSO, or MFA, then re-run."
    )


def configure_probe_logging(artifacts: ArtifactRun) -> logging.Logger:
    probe_logger = logging.getLogger("ispring_cloud")
    probe_logger.setLevel(logging.INFO)
    probe_logger.handlers.clear()
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    console = logging.StreamHandler()
    console.setFormatter(formatter)
    file_handler = logging.FileHandler(artifacts.log_path, encoding="utf-8")
    file_handler.setFormatter(formatter)
    probe_logger.addHandler(console)
    probe_logger.addHandler(file_handler)
    probe_logger.propagate = False
    return probe_logger


def _visible_text(page: Any) -> str:
    chunks: list[str] = []
    for frame in _frames(page):
        try:
            chunks.append(frame.locator("body").inner_text(timeout=2000))
        except Exception:  # noqa: BLE001 — probe must keep going
            continue
    return "\n".join(chunks)


def _frames(page: Any) -> list[Any]:
    try:
        return list(page.frames)
    except Exception:  # noqa: BLE001
        return [page]


def _page_is_open(page: Any) -> bool:
    try:
        return not page.is_closed()
    except Exception:  # noqa: BLE001
        return False


def active_page(context: Any, page: Any | None = None) -> Any:
    if page is not None and _page_is_open(page):
        return page
    open_pages = [candidate for candidate in context.pages if _page_is_open(candidate)]
    if not open_pages:
        raise ISpringCloudProbeError("Chrome closed all tabs unexpectedly.")
    chosen = open_pages[-1]
    logger.info("switched to open tab %s", chosen.url)
    return chosen


def _save_screenshot(page: Any, artifacts: ArtifactRun, stem: str) -> Path:
    path = artifacts.root / artifacts.screenshot_name(stem)
    if not _page_is_open(page):
        logger.warning("skipped screenshot %s because the tab closed", stem)
        return path
    page.screenshot(path=str(path), full_page=True)
    logger.info("saved screenshot %s", path.name)
    return path


def _dump_page(page: Any, artifacts: ArtifactRun, stem: str) -> None:
    if not _page_is_open(page):
        logger.warning("skipped page dump %s because the tab closed", stem)
        return
    try:
        artifacts.write_text(f"{stem}.url.txt", page.url)
    except Exception:  # noqa: BLE001
        pass
    try:
        artifacts.write_text(f"{stem}.text.txt", _visible_text(page)[:20_000])
    except Exception:  # noqa: BLE001
        pass


def dismiss_cookie_dialogs(page: Any) -> bool:
    for frame in _frames(page):
        for name in COOKIE_BUTTON_NAMES:
            button = frame.get_by_role(
                "button", name=re.compile(rf"^{re.escape(name)}$", re.I)
            )
            try:
                if button.count() == 0:
                    continue
                first = button.first
                if first.is_visible():
                    first.click(timeout=2000)
                    logger.info("dismissed cookie or consent dialog (%s)", name)
                    return True
            except Exception:  # noqa: BLE001
                continue
    return False


def _first_visible(locator: Any) -> Any | None:
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


def find_search_box(page: Any) -> Any:
    strategies: tuple[Callable[[Any], Any], ...] = (
        lambda root: root.get_by_role("searchbox"),
        lambda root: root.get_by_placeholder(re.compile(r"search", re.I)),
        lambda root: root.get_by_label(re.compile(r"search", re.I)),
        lambda root: root.get_by_role("textbox", name=re.compile(r"search", re.I)),
        lambda root: root.get_by_role("combobox", name=re.compile(r"search", re.I)),
        lambda root: root.locator(
            "input[type='search'], input[aria-label*='earch' i], "
            "input[placeholder*='earch' i]"
        ),
    )

    def scan() -> Any | None:
        for frame in _frames(page):
            for strategy in strategies:
                found = _first_visible(strategy(frame))
                if found is not None:
                    return found
        return None

    found = scan()
    if found is not None:
        return found
    deadline = time.monotonic() + 20
    while time.monotonic() < deadline:
        for frame in _frames(page):
            toggle = frame.get_by_role("button", name=re.compile(r"search", re.I))
            button = _first_visible(toggle)
            if button is None:
                continue
            try:
                button.click(timeout=ACTION_TIMEOUT_MS)
                logger.info("opened a Search control to reveal the search field")
            except Exception:  # noqa: BLE001
                continue
            found = scan()
            if found is not None:
                return found
        found = scan()
        if found is not None:
            return found
        time.sleep(0.4)
    raise ISpringCloudProbeError(
        "Could not find a search field on the iSpring Cloud page. "
        "A screenshot and page text were saved under the artifacts directory."
    )


def wait_for_visible_text(page: Any, pattern: re.Pattern[str], timeout_ms: int) -> bool:
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        for frame in _frames(page):
            found = _first_visible(frame.get_by_text(pattern))
            if found is not None:
                return True
        time.sleep(0.25)
    return False


def _accessible_name(locator: Any) -> str:
    for attr in ("aria-label", "title"):
        try:
            value = locator.get_attribute(attr)
        except Exception:  # noqa: BLE001
            value = None
        if value:
            return value.strip()
    try:
        return locator.inner_text(timeout=2000).strip()
    except Exception:  # noqa: BLE001
        return ""


def find_named_matches(page: Any, name: str) -> list[Any]:
    pattern = name_pattern(name)
    roles = (
        "link",
        "button",
        "row",
        "listitem",
        "article",
        "option",
        "gridcell",
        "menuitem",
        "heading",
        "tab",
    )
    matches: list[Any] = []
    seen: set[int] = set()
    for frame in _frames(page):
        for role in roles:
            locator = frame.get_by_role(role, name=pattern)
            try:
                count = locator.count()
            except Exception:  # noqa: BLE001
                continue
            for index in range(count):
                item = locator.nth(index)
                try:
                    if not item.is_visible():
                        continue
                    identity = id(item)
                    if identity in seen:
                        continue
                    label = _accessible_name(item)
                    if is_destructive_label(label):
                        logger.info("skipping destructive control named %r", label)
                        continue
                    seen.add(identity)
                    matches.append(item)
                except Exception:  # noqa: BLE001
                    continue
        text_match = _first_visible(frame.get_by_text(pattern))
        if text_match is not None:
            label = _accessible_name(text_match)
            if not is_destructive_label(label):
                matches.append(text_match)
    return matches


def click_maybe_new_tab(locator: Any, context: Any, page: Any) -> Any:
    label = _accessible_name(locator)
    if is_destructive_label(label):
        raise ISpringCloudProbeError(
            f"Refusing to click destructive control: {label!r}"
        )
    pages_before = list(context.pages)
    logger.info("clicking %r", label or "(unnamed control)")
    try:
        with context.expect_page(timeout=4000) as popup_info:
            locator.click(timeout=ACTION_TIMEOUT_MS)
        new_page = popup_info.value
        new_page.wait_for_load_state("domcontentloaded", timeout=NAV_TIMEOUT_MS)
        logger.info("opened a new tab: %s", new_page.url)
        return new_page
    except Exception:  # noqa: BLE001 — no popup is common
        pass
    try:
        page.wait_for_load_state("domcontentloaded", timeout=NAV_TIMEOUT_MS)
    except Exception:  # noqa: BLE001
        pass
    for candidate in context.pages:
        if candidate not in pages_before:
            candidate.wait_for_load_state("domcontentloaded", timeout=NAV_TIMEOUT_MS)
            logger.info("detected a new tab: %s", candidate.url)
            return candidate
    return page


def collect_share_controls(page: Any) -> tuple[str, ...]:
    pattern = re.compile(r"share", re.I)
    names: list[str] = []
    for frame in _frames(page):
        for role in ("button", "link", "menuitem"):
            locator = frame.get_by_role(role, name=pattern)
            try:
                count = locator.count()
            except Exception:  # noqa: BLE001
                continue
            for index in range(count):
                item = locator.nth(index)
                try:
                    if not item.is_visible():
                        continue
                    label = _accessible_name(item) or f"{role}:Share"
                    if label not in names:
                        names.append(label)
                        logger.info("found share control: %s", label)
                except Exception:  # noqa: BLE001
                    continue
        text_hit = _first_visible(frame.get_by_text(pattern))
        if text_hit is not None:
            label = _accessible_name(text_hit)
            if label and label not in names:
                names.append(label)
                logger.info("found share control: %s", label)
    return tuple(names)


def extract_title(page: Any, material_name: str) -> str:
    pattern = name_pattern(material_name)
    for frame in _frames(page):
        heading = _first_visible(frame.get_by_role("heading", name=pattern))
        if heading is not None:
            text = _accessible_name(heading)
            if text:
                return text
        labeled = _first_visible(frame.get_by_text(pattern))
        if labeled is not None:
            text = _accessible_name(labeled)
            if text:
                return text
    for frame in _frames(page):
        heading = _first_visible(frame.get_by_role("heading"))
        if heading is not None:
            text = _accessible_name(heading)
            if text:
                return text
    try:
        title = (page.title() or "").strip()
    except Exception:  # noqa: BLE001
        title = ""
    return title or material_name


def inspect_material(page: Any, material_name: str) -> MaterialInfo:
    url = page.url
    for frame in _frames(page):
        try:
            frame_url = frame.url
        except Exception:  # noqa: BLE001
            continue
        if is_official_ispring_url(frame_url) and extract_material_id(frame_url):
            url = frame_url
            break
    info = MaterialInfo(
        title=extract_title(page, material_name),
        url=url,
        material_id=extract_material_id(url),
        share_controls=collect_share_controls(page),
    )
    logger.info(
        "material title=%r url=%s id=%s share=%s",
        info.title,
        info.url,
        info.material_id,
        ",".join(info.share_controls) or "(none visible)",
    )
    return info


def _open_search_engine_and_pick(page: Any) -> None:
    query_url = search_engine_url(SITE_SEARCH_QUERY)
    logger.info("searching for the official iSpring Cloud site: %s", query_url)
    page.goto(query_url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS)
    dismiss_cookie_dialogs(page)
    links = page.get_by_role("link")
    results: list[tuple[str, str]] = []
    try:
        count = min(links.count(), 40)
    except Exception as exc:  # noqa: BLE001
        raise ISpringCloudProbeError(
            "Could not read search engine results for iSpring Cloud."
        ) from exc
    for index in range(count):
        link = links.nth(index)
        try:
            href = link.get_attribute("href") or ""
            title = (link.inner_text(timeout=1000) or "").strip()
        except Exception:  # noqa: BLE001
            continue
        if href:
            results.append((title, href))
    chosen = pick_official_search_result(results)
    if not chosen:
        raise ISpringCloudProbeError(
            "Search did not return an official iSpring Cloud URL."
        )
    logger.info("opening official search result %s", chosen)
    page.goto(chosen, wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS)


def open_ispring_cloud(page: Any, start_url: str | None) -> None:
    target = start_url or DEFAULT_TEAM_APP_URL
    logger.info("navigating to %s", target)
    try:
        page.goto(target, wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS)
    except Exception as exc:  # noqa: BLE001
        logger.warning("direct navigation failed: %s", exc)
        _open_search_engine_and_pick(page)
        dismiss_cookie_dialogs(page)
        return
    dismiss_cookie_dialogs(page)
    if is_official_ispring_url(page.url):
        return
    logger.info("landed on a non-official host; searching for the official site")
    _open_search_engine_and_pick(page)
    dismiss_cookie_dialogs(page)


def open_cloud_library(context: Any, page: Any) -> Any:
    page = active_page(context, page)
    dismiss_cookie_dialogs(page)
    if is_team_app_url(page.url):
        logger.info("already on the iSpring team app: %s", page.url)
        return page
    cloud_link = _first_visible(
        page.get_by_role("link", name=re.compile(r"ispring cloud", re.I))
    )
    if cloud_link is None:
        cloud_link = _first_visible(
            page.get_by_role("button", name=re.compile(r"ispring cloud", re.I))
        )
    if cloud_link is not None:
        logger.info("opening iSpring Cloud from the account page")
        page = click_maybe_new_tab(cloud_link, context, page)
        page = active_page(context, page)
        dismiss_cookie_dialogs(page)
        if is_team_app_url(page.url):
            return page
    logger.info("navigating to the team app %s", DEFAULT_TEAM_APP_URL)
    try:
        page.goto(DEFAULT_TEAM_APP_URL, wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS)
    except Exception as exc:  # noqa: BLE001
        logger.warning("team app navigation failed: %s", exc)
        page = active_page(context, page)
    return active_page(context, page)


def maybe_wait_for_login(
    page: Any,
    pause: PauseFn | None,
    *,
    email: str | None = None,
    password: str | None = None,
) -> None:
    if not _page_is_open(page):
        logger.info("login tab already closed; continuing")
        return
    text = _visible_text(page)
    if not looks_like_login(page.url, text):
        logger.info("no login form detected; continuing with the current session")
        return
    if email and password:
        submit_login_form(page, email, password)
        return
    logger.info(
        "login required. complete login, SSO, or MFA in the Chrome window. "
        "credentials are not stored in the repository."
    )
    print("\nLogin is required in the Chrome window.")
    print("Complete login, SSO, or MFA there. Do not close the browser.")
    print("This script does not store credentials in the repository.\n")
    if pause is None:
        raise ISpringCloudProbeError(
            "Login is required, but no interactive pause is available."
        )
    pause("Press Enter here after login is complete: ")
    try:
        page.wait_for_load_state("domcontentloaded", timeout=NAV_TIMEOUT_MS)
    except Exception:  # noqa: BLE001
        pass
    dismiss_cookie_dialogs(page)
    logger.info("resumed after manual login confirmation")


def wait_for_app_ready(page: Any, timeout_ms: int = 45_000) -> None:
    logger.info("waiting for the team app UI to finish loading")
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        if not _page_is_open(page):
            return
        try:
            text = _visible_text(page)
        except Exception:  # noqa: BLE001
            time.sleep(0.4)
            continue
        if looks_like_login(page.url, text):
            time.sleep(0.4)
            continue
        if re.search(
            r"demoacademy|no results found|\btitle\b|results:|search\.\.\.",
            text,
            re.I,
        ):
            logger.info("app UI is visible (%s characters)", len(text))
            return
        time.sleep(0.4)
    logger.warning("app UI still looks empty after waiting; continuing")


def search_library(page: Any, query: str) -> None:
    logger.info("searching the library for %r", query)
    box = find_search_box(page)
    box.click(timeout=ACTION_TIMEOUT_MS)
    box.fill(query, timeout=ACTION_TIMEOUT_MS)
    submit = None
    for frame in _frames(page):
        submit = _first_visible(
            frame.get_by_role("button", name=re.compile(r"^search$", re.I))
        )
        if submit is not None:
            break
    if submit is not None:
        submit.click(timeout=ACTION_TIMEOUT_MS)
    else:
        box.press("Enter")
    wait_for_visible_text(page, name_pattern(query), SEARCH_TIMEOUT_MS)


def _label_word_set(label: str) -> set[str]:
    return {word for word in normalize_text(label).split() if word}


def find_project_matches(page: Any, query: str) -> list[Any]:
    query_words = _label_word_set(query)
    if not query_words:
        return []
    exact: list[Any] = []
    partial: list[tuple[int, Any]] = []
    roles = ("link", "button", "treeitem", "listitem", "menuitem")
    for frame in _frames(page):
        for role in roles:
            locator = frame.get_by_role(role)
            try:
                count = locator.count()
            except Exception:  # noqa: BLE001
                continue
            for index in range(count):
                item = locator.nth(index)
                try:
                    if not item.is_visible():
                        continue
                    label = _accessible_name(item)
                    words = _label_word_set(label)
                    if not words or is_destructive_label(label):
                        continue
                    if words == query_words:
                        exact.append(item)
                    elif words <= query_words or query_words <= words:
                        partial.append((len(words & query_words), item))
                except Exception:  # noqa: BLE001
                    continue
    partial.sort(key=lambda row: row[0], reverse=True)
    return exact + [item for _, item in partial]


def open_first_material(page: Any, context: Any, query: str, material_name: str) -> Any:
    logger.info("looking for material %r", material_name)
    if wait_for_visible_text(page, name_pattern(material_name), 5_000):
        matches = find_named_matches(page, material_name)
        if matches:
            return click_maybe_new_tab(matches[0], context, page)
    logger.info(
        "%r was not visible in the first result list; trying a folder/project named %r",
        material_name,
        query,
    )
    try:
        box = find_search_box(page)
        box.fill("")
        box.press("Escape")
        logger.info("cleared the search box so project folders are visible")
        wait_for_app_ready(page, timeout_ms=15_000)
    except Exception:  # noqa: BLE001
        pass
    project_hits = find_project_matches(page, query)
    if not project_hits:
        for frame in _frames(page):
            text_hit = _first_visible(
                frame.get_by_text(re.compile(r"ppt\s*migration(\s+new)?", re.I))
            )
            if text_hit is not None:
                project_hits.append(text_hit)
                break
    for project in project_hits:
        page = click_maybe_new_tab(project, context, page)
        dismiss_cookie_dialogs(page)
        wait_for_app_ready(page, timeout_ms=20_000)
        if wait_for_visible_text(page, name_pattern(material_name), SEARCH_TIMEOUT_MS):
            matches = find_named_matches(page, material_name)
            if matches:
                return click_maybe_new_tab(matches[0], context, page)
    folders = find_named_matches(page, query)
    if folders:
        page = click_maybe_new_tab(folders[0], context, page)
        dismiss_cookie_dialogs(page)
        if wait_for_visible_text(page, name_pattern(material_name), SEARCH_TIMEOUT_MS):
            matches = find_named_matches(page, material_name)
            if matches:
                return click_maybe_new_tab(matches[0], context, page)
    raise ISpringCloudProbeError(
        f"Could not find a material matching {material_name!r} after searching "
        f"for {query!r}."
    )


def launch_persistent_chrome(playwright: Any, profile_dir: Path, *, headless: bool) -> Any:
    profile_dir.mkdir(parents=True, exist_ok=True)
    logger.info("launching Google Chrome with persistent profile %s", profile_dir)
    try:
        return playwright.chromium.launch_persistent_context(
            str(profile_dir),
            channel="chrome",
            headless=headless,
            viewport={"width": 1440, "height": 900},
        )
    except Exception as exc:  # noqa: BLE001
        message = str(exc)
        raise ISpringCloudProbeError(
            "Could not launch Google Chrome via Playwright "
            f"(channel='chrome'): {message}. Install Google Chrome and close any "
            "Chrome window that is already using this profile directory."
        ) from exc


def run_cloud_probe(
    *,
    artifacts: ArtifactRun,
    profile_dir: Path,
    start_url: str | None = None,
    search_query: str = DEFAULT_SEARCH_QUERY,
    material_name: str = DEFAULT_MATERIAL_NAME,
    headless: bool = False,
    pause: PauseFn | None = input,
    email: str | None = None,
    password: str | None = None,
) -> MaterialInfo:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise ISpringCloudProbeError(
            "playwright is not installed. Run: pip install -r requirements.txt"
        ) from exc

    configure_probe_logging(artifacts)
    logger.info("starting read-only iSpring Cloud probe")
    logger.info("artifacts directory: %s", artifacts.root)
    logger.info("this probe does not upload, delete, publish, or edit materials")

    with sync_playwright() as playwright:
        context = launch_persistent_chrome(
            playwright, profile_dir, headless=headless
        )
        try:
            context.tracing.start(screenshots=True, snapshots=True, sources=True)
            page = context.pages[0] if context.pages else context.new_page()
            _save_screenshot(page, artifacts, "01-browser-open")
            open_ispring_cloud(page, start_url)
            page = active_page(context, page)
            _save_screenshot(page, artifacts, "02-ispring-cloud-loaded")
            _dump_page(page, artifacts, "02-ispring-cloud-loaded")
            try:
                maybe_wait_for_login(page, pause, email=email, password=password)
            except ISpringCloudProbeError:
                page = active_page(context, page)
                _save_screenshot(page, artifacts, "03-login-failed")
                _dump_page(page, artifacts, "03-login-failed")
                raise
            page = open_cloud_library(context, page)
            if looks_like_login(page.url, _visible_text(page)):
                logger.info("iSpring Cloud app still shows its own login form")
                try:
                    maybe_wait_for_login(
                        page, pause, email=email, password=password
                    )
                except ISpringCloudProbeError:
                    page = active_page(context, page)
                    _save_screenshot(page, artifacts, "03-login-failed")
                    _dump_page(page, artifacts, "03-login-failed")
                    raise
                page = active_page(context, page)
            if looks_like_login(page.url, _visible_text(page)):
                raise ISpringCloudProbeError(
                    "Still on a login page after the pause. Re-run after finishing login."
                )
            wait_for_app_ready(page)
            _save_screenshot(page, artifacts, "03-login-complete")
            _dump_page(page, artifacts, "03-login-complete")
            already_open = wait_for_visible_text(
                page, name_pattern(material_name), 8_000
            )
            if already_open:
                logger.info(
                    "%r is already visible from the start URL; skipping keyword search",
                    material_name,
                )
                _save_screenshot(page, artifacts, "04-search-new-ppt-migration")
                _dump_page(page, artifacts, "04-search-new-ppt-migration")
            else:
                search_library(page, search_query)
                _save_screenshot(page, artifacts, "04-search-new-ppt-migration")
                _dump_page(page, artifacts, "04-search-new-ppt-migration")
            _save_screenshot(page, artifacts, "05-found-demo-academy")
            try:
                page = open_first_material(page, context, search_query, material_name)
            except ISpringCloudProbeError:
                _dump_page(page, artifacts, "05-material-not-found")
                raise
            dismiss_cookie_dialogs(page)
            try:
                page.wait_for_load_state("domcontentloaded", timeout=NAV_TIMEOUT_MS)
            except Exception:  # noqa: BLE001
                pass
            _save_screenshot(page, artifacts, "06-material-opened")
            _dump_page(page, artifacts, "06-material-opened")
            info = inspect_material(page, material_name)
            artifacts.write_text(
                "material.json",
                json.dumps(
                    {
                        "title": info.title,
                        "url": info.url,
                        "material_id": info.material_id,
                        "share_controls": list(info.share_controls),
                    },
                    indent=2,
                )
                + "\n",
            )
            return info
        finally:
            try:
                context.tracing.stop(path=str(artifacts.root / "trace.zip"))
                logger.info("saved Playwright trace to trace.zip")
            except Exception as exc:  # noqa: BLE001
                logger.warning("could not save Playwright trace: %s", exc)
            context.close()
            logger.info("closed Chrome; persistent profile kept at %s", profile_dir)
