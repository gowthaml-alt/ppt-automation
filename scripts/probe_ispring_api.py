"""Second iSpring probe: can the add-in be driven without the mouse?

Read-only. Never publishes, never launches an iSpring window. Run on the
Windows machine that has iSpring Suite installed, then attach the report.

    python scripts\\probe_ispring_api.py

Writes LOG_ROOT/ispring-api-probe.json (default logs/ispring-api-probe.json).
"""

from __future__ import annotations

import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

SUITE_PROGID_HINT = "ispring"

# Names worth asking the add-in about. Asking is a dispatch-ID lookup only;
# nothing is called.
CANDIDATE_MEMBERS = [
    "Publish",
    "PublishPresentation",
    "PublishToHTML",
    "PublishToHTML5",
    "PublishHTML5",
    "PublishToWeb",
    "PublishToFolder",
    "Build",
    "BuildPresentation",
    "Convert",
    "Export",
    "ExportToHTML5",
    "Generate",
    "Run",
    "Execute",
    "Automation",
    "Application",
    "Version",
    "Settings",
    "Options",
    "Presentation",
    "Project",
]

CLI_CANDIDATES = [
    r"bin32\ispringlauncher.exe",
    r"bin32\iSpringSvr.exe",
    r"bin32\ispringuploader.exe",
    r"bin32\ispringpreview.exe",
    r"bin32\ComLauncher.exe",
    r"bin32\comcefview.exe",
    r"bin\activation.exe",
]

SWITCH_RE = re.compile(
    rb"(?<![A-Za-z0-9:.\\/_-])(?:--|/)[A-Za-z][A-Za-z0-9_-]{2,24}"
)


def _safe(fn, default=None):
    try:
        return fn()
    except Exception as exc:  # noqa: BLE001 - a probe must keep going
        return {"error": f"{type(exc).__name__}: {exc}"}


def _describe_type_info(dispatch) -> dict:
    """List the methods and properties the add-in object declares, if any."""
    oleobj = getattr(dispatch, "_oleobj_", None)
    if oleobj is None:
        return {"available": False, "reason": "object has no _oleobj_"}
    try:
        if int(oleobj.GetTypeInfoCount()) == 0:
            return {"available": False, "reason": "GetTypeInfoCount() == 0"}
    except Exception as exc:  # noqa: BLE001
        return {"available": False, "reason": f"GetTypeInfoCount failed: {exc}"}
    try:
        info = oleobj.GetTypeInfo()
        attr = info.GetTypeAttr()
    except Exception as exc:  # noqa: BLE001
        return {"available": False, "reason": f"GetTypeInfo failed: {exc}"}

    funcs = []
    for index in range(attr.cFuncs):
        try:
            desc = info.GetFuncDesc(index)
            names = info.GetNames(desc.memid)
        except Exception:  # noqa: BLE001
            continue
        funcs.append(
            {
                "name": names[0] if names else "",
                "args": list(names[1:]),
                "arg_count": int(desc.cParams),
                "invoke_kind": int(desc.invkind),
            }
        )
    variables = []
    for index in range(attr.cVars):
        try:
            desc = info.GetVarDesc(index)
            names = info.GetNames(desc.memid)
        except Exception:  # noqa: BLE001
            continue
        if names:
            variables.append(names[0])
    return {
        "available": True,
        "function_count": int(attr.cFuncs),
        "functions": funcs,
        "variables": variables,
    }


def _lookup_members(dispatch) -> dict:
    """Ask the object whether each candidate name exists. No calls are made."""
    oleobj = getattr(dispatch, "_oleobj_", None)
    if oleobj is None:
        return {"error": "object has no _oleobj_"}
    found = {}
    for name in CANDIDATE_MEMBERS:
        try:
            found[name] = int(oleobj.GetIDsOfNames(name))
        except Exception:  # noqa: BLE001
            continue
    return found


def probe_addin_objects() -> list[dict]:
    import win32com.client  # type: ignore

    app = win32com.client.Dispatch("PowerPoint.Application")
    rows = []
    for addin in app.COMAddIns:
        prog_id = str(getattr(addin, "ProgId", ""))
        if SUITE_PROGID_HINT not in prog_id.lower():
            continue
        row = {
            "prog_id": prog_id,
            "description": str(getattr(addin, "Description", "")),
            "connect": bool(getattr(addin, "Connect", False)),
        }
        try:
            obj = addin.Object
        except Exception as exc:  # noqa: BLE001
            row["object"] = None
            row["object_error"] = f"{type(exc).__name__}: {exc}"
            rows.append(row)
            continue
        if obj is None:
            row["object"] = None
            row["object_error"] = "addin.Object is None (no automation service)"
            rows.append(row)
            continue
        row["object"] = repr(obj)[:400]
        row["type_info"] = _describe_type_info(obj)
        row["members_found"] = _lookup_members(obj)
        rows.append(row)
    return rows


def probe_powerpoint_addins() -> list[dict]:
    """Classic .ppa/.ppam add-ins can expose callable macros via Application.Run."""
    import win32com.client  # type: ignore

    app = win32com.client.Dispatch("PowerPoint.Application")
    rows = []
    for addin in app.AddIns:
        rows.append(
            {
                "name": str(getattr(addin, "Name", "")),
                "full_name": str(getattr(addin, "FullName", "")),
                "loaded": bool(getattr(addin, "Loaded", False)),
                "registered": bool(getattr(addin, "Registered", False)),
            }
        )
    return rows


def probe_registry() -> dict:
    try:
        import winreg  # type: ignore
    except ImportError:
        return {"error": "winreg unavailable"}

    hives = (
        (winreg.HKEY_CURRENT_USER, "HKCU"),
        (winreg.HKEY_LOCAL_MACHINE, "HKLM"),
    )
    bases = [r"SOFTWARE", r"SOFTWARE\WOW6432Node"]
    matches: list[str] = []
    values: dict[str, dict] = {}

    for hive, hive_name in hives:
        for base in bases:
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
                if "ispring" not in name.lower():
                    continue
                key_path = f"{base}\\{name}"
                matches.append(f"{hive_name}\\{key_path}")
                values[f"{hive_name}\\{key_path}"] = _read_key_tree(
                    winreg, hive, key_path, depth=3
                )
            handle.Close()
    return {"keys": matches, "values": values}


def _read_key_tree(winreg, hive, key_path: str, depth: int) -> dict:
    """Read value names and short values. Publish folders often live here."""
    out: dict = {}
    if depth <= 0:
        return {"truncated": True}
    try:
        handle = winreg.OpenKey(hive, key_path)
    except OSError as exc:
        return {"error": str(exc)}
    index = 0
    while True:
        try:
            name, value, _kind = winreg.EnumValue(handle, index)
        except OSError:
            break
        index += 1
        text = str(value)
        out[name or "(default)"] = text[:300]
    index = 0
    children: dict = {}
    while True:
        try:
            sub = winreg.EnumKey(handle, index)
        except OSError:
            break
        index += 1
        children[sub] = _read_key_tree(winreg, hive, f"{key_path}\\{sub}", depth - 1)
    handle.Close()
    if children:
        out["_subkeys"] = children
    return out


def probe_cli_switches() -> list[dict]:
    """Scan candidate binaries for switch-looking strings. Nothing is executed."""
    roots = [
        os.environ.get("ISPRING_INSTALL_PATH", ""),
        r"C:\Program Files\iSpring\Suite 11",
        r"C:\Program Files\iSpring\Free 11",
    ]
    rows = []
    for root in roots:
        if not root:
            continue
        base = Path(root)
        if not base.exists():
            continue
        for relative in CLI_CANDIDATES:
            target = base / relative
            if not target.is_file():
                continue
            try:
                blob = target.read_bytes()
            except OSError as exc:
                rows.append({"path": str(target), "error": str(exc)})
                continue
            # Nulls become spaces so UTF-16 text is readable and 8-bit strings
            # stay separated instead of running together.
            flat = blob.replace(b"\x00", b" ")
            hits = sorted({m.decode("ascii", "ignore") for m in SWITCH_RE.findall(flat)})
            rows.append(
                {
                    "path": str(target),
                    "size": target.stat().st_size,
                    "switch_like_strings": hits[:80],
                    "switch_like_count": len(hits),
                }
            )
    return rows


def probe_recent_output(limit: int = 40) -> list[str]:
    """Find index.html files iSpring may have written during manual publishes."""
    roots = [
        Path.home() / "Documents",
        Path.home() / "Desktop",
        Path.home() / "Downloads",
    ]
    found: list[str] = []
    for root in roots:
        if not root.exists():
            continue
        try:
            for item in root.rglob("index.html"):
                found.append(str(item))
                if len(found) >= limit:
                    return found
        except OSError:
            continue
    return found


def main() -> int:
    if sys.platform != "win32":
        print(
            "This probe needs Windows with PowerPoint and iSpring installed.",
            file=sys.stderr,
        )
        return 2

    report = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "addin_objects": _safe(probe_addin_objects, []),
        "powerpoint_addins": _safe(probe_powerpoint_addins, []),
        "registry": _safe(probe_registry, {}),
        "cli_candidates": _safe(probe_cli_switches, []),
        "index_html_found": _safe(probe_recent_output, []),
        "still_needed_by_hand": [
            "Publish one deck manually with iSpring, then say which folder "
            "index.html landed in and whether any dialog needed an answer.",
            "Roughly how many minutes that publish took.",
        ],
    }
    log_root = Path(os.environ.get("LOG_ROOT", "logs"))
    log_root.mkdir(parents=True, exist_ok=True)
    out = log_root / "ispring-api-probe.json"
    out.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(f"Wrote {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
