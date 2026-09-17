"""Keep the iSpring add-in loaded in PowerPoint.

Office watches for add-ins that were loaded when PowerPoint died. Any forced
close — which this worker sometimes has to do — looks like a crash, so Office
moves the add-in into Disabled Items and the iSpring ribbon tab stops
appearing. A job that starts in that state cannot publish at all.

Two registry facts decide it:

``LoadBehavior``   3 means load at startup; Office writes 2 when it demotes it
``DisabledItems``  Office's own list of things it has switched off

Both are under HKCU (per user) and HKLM (per machine). HKCU is the one Office
demotes, and the one this can always write.
"""

from __future__ import annotations

import logging
import sys
from dataclasses import dataclass, field

logger = logging.getLogger(__name__)

STAGE = "powerpoint_open"
ADDIN_HINT = "ispring"
LOAD_AT_STARTUP = 3

# Office's own policy list. An add-in listed here as "1" is always enabled and
# Office will not disable it after a crash — which is what stops the ribbon tab
# coming and going. RestrictToList is deliberately never written: that one
# blocks every add-in not in the list.
POLICY_BASE = r"SOFTWARE\Policies\Microsoft\Office"
POLICY_ALWAYS_ENABLED = "1"

ADDIN_BASES = (
    r"SOFTWARE\Microsoft\Office\PowerPoint\Addins",
    r"SOFTWARE\WOW6432Node\Microsoft\Office\PowerPoint\Addins",
)

LOAD_BEHAVIOUR_MEANING = {
    0: "never load",
    1: "load once, then off",
    2: "disabled by Office",
    3: "load at startup",
}


@dataclass
class AddinState:
    registered: list[str] = field(default_factory=list)
    demoted: list[str] = field(default_factory=list)
    disabled_items: list[tuple[str, str]] = field(default_factory=list)
    checked: bool = True

    @property
    def healthy(self) -> bool:
        if not self.checked:
            return True  # nothing to say off Windows
        return bool(self.registered) and not self.demoted and not self.disabled_items


def _winreg():
    import winreg  # type: ignore

    return winreg


def _addin_keys() -> list[tuple[int, str, str]]:
    winreg = _winreg()
    found = []
    for hive, hive_name in (
        (winreg.HKEY_CURRENT_USER, "HKCU"),
        (winreg.HKEY_LOCAL_MACHINE, "HKLM"),
    ):
        for base in ADDIN_BASES:
            try:
                handle = winreg.OpenKey(hive, base)
            except OSError:
                continue
            index = 0
            while True:
                try:
                    name = winreg.EnumKey(handle, index)
                except OSError:
                    break
                index += 1
                if ADDIN_HINT in name.lower():
                    found.append((hive, f"{base}\\{name}", f"{hive_name}\\{base}\\{name}"))
            handle.Close()
    return found


def _load_behavior(hive: int, path: str) -> int | None:
    winreg = _winreg()
    try:
        with winreg.OpenKey(hive, path) as handle:
            value, _kind = winreg.QueryValueEx(handle, "LoadBehavior")
            return int(value)
    except OSError:
        return None


def _set_load_behavior(hive: int, path: str) -> bool:
    winreg = _winreg()
    try:
        with winreg.OpenKey(hive, path, 0, winreg.KEY_SET_VALUE) as handle:
            winreg.SetValueEx(handle, "LoadBehavior", 0, winreg.REG_DWORD, LOAD_AT_STARTUP)
        return True
    except OSError as exc:
        logger.warning(
            "could not set LoadBehavior; an elevated prompt may be needed",
            extra={"stage": STAGE, "key": path, "error": str(exc)},
        )
        return False


def _disabled_items() -> list[tuple[int, str, str]]:
    winreg = _winreg()
    hive = winreg.HKEY_CURRENT_USER
    base = r"SOFTWARE\Microsoft\Office"
    found: list[tuple[int, str, str]] = []
    try:
        office = winreg.OpenKey(hive, base)
    except OSError:
        return found
    versions = []
    index = 0
    while True:
        try:
            versions.append(winreg.EnumKey(office, index))
        except OSError:
            break
        index += 1
    office.Close()
    for version in versions:
        path = f"{base}\\{version}\\PowerPoint\\Resiliency\\DisabledItems"
        try:
            handle = winreg.OpenKey(hive, path)
        except OSError:
            continue
        value_index = 0
        while True:
            try:
                name, _value, _kind = winreg.EnumValue(handle, value_index)
            except OSError:
                break
            value_index += 1
            found.append((hive, path, name))
        handle.Close()
    return found


def _clear_disabled(hive: int, path: str, name: str) -> bool:
    winreg = _winreg()
    try:
        with winreg.OpenKey(hive, path, 0, winreg.KEY_SET_VALUE) as handle:
            winreg.DeleteValue(handle, name)
        return True
    except OSError as exc:
        logger.warning(
            "could not clear a disabled item",
            extra={"stage": STAGE, "key": path, "value": name, "error": str(exc)},
        )
        return False


def check_addin() -> AddinState:
    """What state the add-in registration is in. Reads only."""
    if sys.platform != "win32":
        return AddinState(checked=False)
    state = AddinState()
    for hive, path, label in _addin_keys():
        state.registered.append(label)
        behaviour = _load_behavior(hive, path)
        if behaviour != LOAD_AT_STARTUP:
            state.demoted.append(
                f"{label} (LoadBehavior={behaviour}: "
                f"{LOAD_BEHAVIOUR_MEANING.get(behaviour, 'unknown')})"
            )
    for _hive, path, name in _disabled_items():
        state.disabled_items.append((path, name))
    return state


def ensure_addin_enabled() -> bool:
    """Check the add-in before a job, and put it back if Office switched it off.

    Called with PowerPoint closed: Office rewrites these keys when it exits, so
    fixing them under a running PowerPoint achieves nothing.

    Returns True when the add-in should load. False means it is not registered
    at all, which needs an iSpring repair install — not something to fix here.
    """
    state = check_addin()
    if not state.checked:
        return True
    if not state.registered:
        logger.error(
            "the iSpring add-in is not registered with PowerPoint at all; "
            "repair the iSpring Suite installation",
            extra={"stage": STAGE},
        )
        return False
    harden_addin()  # cheap, idempotent, and stops the next crash disabling it

    if state.healthy:
        logger.info(
            "iSpring add-in is enabled",
            extra={"stage": STAGE, "keys": len(state.registered)},
        )
        return True

    logger.warning(
        "Office had switched the iSpring add-in off; turning it back on",
        extra={
            "stage": STAGE,
            "demoted": state.demoted,
            "disabled_items": [name for _path, name in state.disabled_items],
        },
    )
    for hive, path, label in _addin_keys():
        if _load_behavior(hive, path) != LOAD_AT_STARTUP:
            if _set_load_behavior(hive, path):
                logger.info(
                    "set LoadBehavior to 3", extra={"stage": STAGE, "key": label}
                )
    for hive, path, name in _disabled_items():
        if _clear_disabled(hive, path, name):
            logger.info(
                "cleared a disabled item", extra={"stage": STAGE, "value": name}
            )

    after = check_addin()
    if after.healthy:
        logger.info("iSpring add-in re-enabled", extra={"stage": STAGE})
        return True
    logger.error(
        "the iSpring add-in is still switched off; run "
        "scripts/fix_ispring_addin.py --fix from an elevated prompt",
        extra={"stage": STAGE, "demoted": after.demoted},
    )
    return False


def connect_addin(app) -> bool:
    """Switch the iSpring add-in on inside a PowerPoint that is already running.

    PowerPoint keeps its loaded add-ins in the COMAddIns collection, and
    setting Connect to True loads one there and then. This is worth trying
    before restarting PowerPoint: it takes a second and usually brings the
    ribbon tab back on its own.

    Returns True when an iSpring add-in is connected afterwards.
    """
    if app is None:
        return False
    try:
        addins = app.COMAddIns
        count = int(addins.Count)
    except Exception as exc:  # noqa: BLE001
        logger.warning(
            "could not read PowerPoint's add-in list",
            extra={"stage": STAGE, "error": str(exc)},
        )
        return False

    connected = False
    for index in range(1, count + 1):
        try:
            item = addins.Item(index)
            prog_id = str(getattr(item, "ProgID", "") or "")
        except Exception:  # noqa: BLE001
            continue
        if ADDIN_HINT not in prog_id.lower():
            continue
        try:
            if bool(getattr(item, "Connect", False)):
                logger.info(
                    "iSpring add-in already connected",
                    extra={"stage": STAGE, "prog_id": prog_id},
                )
                connected = True
                continue
            item.Connect = True
            connected = bool(getattr(item, "Connect", False)) or connected
            logger.info(
                "switched the iSpring add-in on in the running PowerPoint",
                extra={"stage": STAGE, "prog_id": prog_id},
            )
        except Exception as exc:  # noqa: BLE001
            logger.warning(
                "could not connect the iSpring add-in through PowerPoint",
                extra={"stage": STAGE, "prog_id": prog_id, "error": str(exc)},
            )
    if not connected:
        logger.warning(
            "no iSpring add-in in PowerPoint's add-in list",
            extra={"stage": STAGE, "addins_seen": count},
        )
    return connected


def _office_versions() -> list[str]:
    """Office version keys present for PowerPoint, newest first."""
    winreg = _winreg()
    versions = []
    try:
        handle = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Office")
    except OSError:
        return ["16.0"]
    index = 0
    while True:
        try:
            name = winreg.EnumKey(handle, index)
        except OSError:
            break
        index += 1
        if not name[0].isdigit():
            continue
        try:
            winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, rf"SOFTWARE\Microsoft\Office\{name}\PowerPoint"
            ).Close()
        except OSError:
            continue
        versions.append(name)
    return sorted(versions, reverse=True) or ["16.0"]


def addin_prog_ids() -> list[str]:
    """The ProgIDs of the registered iSpring add-ins."""
    return [path.rsplit("\\", 1)[-1] for _hive, path, _label in _addin_keys()]


def harden_addin() -> bool:
    """Tell Office never to disable the iSpring add-in.

    Without this, every forced close of PowerPoint can make Office demote the
    add-in, and the ribbon tab disappears until someone re-enables it by hand.
    Listing the add-in in the policy AddinList as "1" makes that permanent:
    Office keeps it enabled no matter what it thinks happened.
    """
    if sys.platform != "win32":
        return False
    winreg = _winreg()
    prog_ids = addin_prog_ids()
    if not prog_ids:
        return False
    wrote = False
    for version in _office_versions():
        path = rf"{POLICY_BASE}\{version}\PowerPoint\Resiliency\AddinList"
        try:
            handle = winreg.CreateKeyEx(
                winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_SET_VALUE | winreg.KEY_READ
            )
        except OSError as exc:
            logger.warning(
                "could not open the add-in policy key",
                extra={"stage": STAGE, "key": path, "error": str(exc)},
            )
            continue
        with handle:
            for prog_id in prog_ids:
                try:
                    existing, _kind = winreg.QueryValueEx(handle, prog_id)
                except OSError:
                    existing = None
                if str(existing) == POLICY_ALWAYS_ENABLED:
                    continue
                try:
                    winreg.SetValueEx(
                        handle, prog_id, 0, winreg.REG_SZ, POLICY_ALWAYS_ENABLED
                    )
                    wrote = True
                    logger.info(
                        "add-in pinned to always enabled",
                        extra={"stage": STAGE, "prog_id": prog_id, "office": version},
                    )
                except OSError as exc:
                    logger.warning(
                        "could not pin the add-in",
                        extra={"stage": STAGE, "prog_id": prog_id, "error": str(exc)},
                    )
    return wrote
