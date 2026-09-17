"""Bring the iSpring add-in back after Office disabled it.

Office watches for add-ins that were loaded when PowerPoint died. Killing
PowerPoint — which this worker does when an instance stops answering — looks
exactly like a crash, so Office moves the add-in to Disabled Items and the
ribbon tab stops appearing.

    python scripts\\fix_ispring_addin.py            show what is wrong
    python scripts\\fix_ispring_addin.py --fix      put it right

Fixing does two things: clears Office's disabled-items list for PowerPoint,
and sets the add-in's LoadBehavior back to 3 (load at startup). Close
PowerPoint first; it rewrites these keys when it exits.
"""

from __future__ import annotations

import argparse
import sys

ADDIN_HINT = "ispring"
LOAD_AT_STARTUP = 3


def _winreg():
    import winreg  # type: ignore

    return winreg


def addin_keys() -> list[tuple[int, str, str]]:
    """Every registered PowerPoint add-in key that mentions iSpring."""
    winreg = _winreg()
    found = []
    for hive, hive_name in (
        (winreg.HKEY_CURRENT_USER, "HKCU"),
        (winreg.HKEY_LOCAL_MACHINE, "HKLM"),
    ):
        for base in (
            r"SOFTWARE\Microsoft\Office\PowerPoint\Addins",
            r"SOFTWARE\WOW6432Node\Microsoft\Office\PowerPoint\Addins",
        ):
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


def load_behavior(hive: int, path: str) -> int | None:
    winreg = _winreg()
    try:
        with winreg.OpenKey(hive, path) as handle:
            value, _kind = winreg.QueryValueEx(handle, "LoadBehavior")
            return int(value)
    except OSError:
        return None


def set_load_behavior(hive: int, path: str) -> bool:
    winreg = _winreg()
    try:
        with winreg.OpenKey(hive, path, 0, winreg.KEY_SET_VALUE) as handle:
            winreg.SetValueEx(handle, "LoadBehavior", 0, winreg.REG_DWORD, LOAD_AT_STARTUP)
        return True
    except OSError as exc:
        print(f"  could not write LoadBehavior ({exc}); try an elevated prompt")
        return False


def disabled_items() -> list[tuple[int, str, str]]:
    """Office's list of add-ins it decided to switch off."""
    winreg = _winreg()
    found = []
    hive = winreg.HKEY_CURRENT_USER
    base = r"SOFTWARE\Microsoft\Office"
    try:
        office = winreg.OpenKey(hive, base)
    except OSError:
        return found
    index = 0
    versions = []
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


def clear_disabled(hive: int, path: str, name: str) -> bool:
    winreg = _winreg()
    try:
        with winreg.OpenKey(hive, path, 0, winreg.KEY_SET_VALUE) as handle:
            winreg.DeleteValue(handle, name)
        return True
    except OSError as exc:
        print(f"  could not clear the disabled entry ({exc})")
        return False


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Check, and optionally repair, the iSpring PowerPoint add-in."
    )
    parser.add_argument("--fix", action="store_true", help="Apply the repairs")
    args = parser.parse_args(argv)

    if sys.platform != "win32":
        print("This runs on the Windows machine.", file=sys.stderr)
        return 2

    problems = 0

    print("iSpring add-in registrations")
    keys = addin_keys()
    if not keys:
        print("  none found — iSpring is not registered with PowerPoint at all.")
        print("  Repair the iSpring Suite installation from Windows Settings.")
        problems += 1
    for hive, path, label in keys:
        current = load_behavior(hive, path)
        state = {0: "never load", 1: "load once", 2: "disabled by Office", 3: "load at startup"}
        print(f"  {label}: LoadBehavior={current} ({state.get(current, 'unknown')})")
        if current == LOAD_AT_STARTUP:
            continue
        problems += 1
        if args.fix and set_load_behavior(hive, path):
            print("    set to 3 (load at startup)")

    print("\nOffice disabled items for PowerPoint")
    disabled = disabled_items()
    if not disabled:
        print("  none — Office has not switched anything off.")
    for hive, path, name in disabled:
        print(f"  {path}\\{name}")
        problems += 1
        if args.fix and clear_disabled(hive, path, name):
            print("    cleared")

    if problems and not args.fix:
        print(f"\n{problems} thing(s) to fix. Close PowerPoint, then run with --fix")
        return 1
    if problems:
        print("\nDone. Start PowerPoint and check the iSpring Suite 11 tab is back.")
        return 0
    print("\nNothing to fix.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
