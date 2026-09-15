"""Gather facts about the installed iSpring / PowerPoint add-in.

Run this on the Windows Cloud PC, then attach the report before implementing
a real ISPRING_ADAPTER. This script never publishes a presentation.
"""

from __future__ import annotations

import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path


def _safe(fn, default=None):
    try:
        return fn()
    except Exception as exc:  # noqa: BLE001 — probe must keep going
        return {"error": f"{type(exc).__name__}: {exc}"}


def probe_com_addins() -> list[dict]:
    import win32com.client  # type: ignore

    app = win32com.client.Dispatch("PowerPoint.Application")
    rows = []
    for addin in app.COMAddIns:
        rows.append(
            {
                "prog_id": str(getattr(addin, "ProgId", "")),
                "description": str(getattr(addin, "Description", "")),
                "connect": bool(getattr(addin, "Connect", False)),
                "guid": str(getattr(addin, "Guid", "")),
            }
        )
    return rows


def probe_install_paths() -> list[dict]:
    candidates = []
    configured = os.environ.get("ISPRING_INSTALL_PATH", "")
    roots = [
        configured,
        r"C:\Program Files\iSpring",
        r"C:\Program Files (x86)\iSpring",
    ]
    for root in roots:
        if not root:
            continue
        path = Path(root)
        if not path.exists():
            candidates.append({"path": root, "exists": False})
            continue
        executables = [
            {"path": str(item), "name": item.name}
            for item in path.rglob("*")
            if item.is_file() and item.suffix.lower() in {".exe", ".dll"}
        ]
        candidates.append(
            {"path": root, "exists": True, "binaries_sample": executables[:50]}
        )
    return candidates


def probe_registry() -> list[str]:
    try:
        import winreg  # type: ignore
    except ImportError:
        return ["winreg unavailable"]
    keys = [
        r"SOFTWARE\iSpring",
        r"SOFTWARE\WOW6432Node\iSpring",
    ]
    found = []
    for hive, hive_name in (
        (winreg.HKEY_LOCAL_MACHINE, "HKLM"),
        (winreg.HKEY_CURRENT_USER, "HKCU"),
    ):
        for key in keys:
            try:
                handle = winreg.OpenKey(hive, key)
            except OSError:
                continue
            found.append(f"{hive_name}\\{key}")
            handle.Close()
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
        "com_addins": _safe(probe_com_addins, []),
        "install_paths": _safe(probe_install_paths, []),
        "registry_keys": _safe(probe_registry, []),
        "needed_to_finish_integration": [
            "Which COMAddIn ProgID is iSpring, and is Connect true?",
            "Is there a callable Application.Run entry or only ribbon UI?",
            "Where does a manual publish write index.html?",
            "How long does a typical publish take (sets ISPRING_PUBLISH_TIMEOUT_SECONDS)?",
        ],
    }
    log_root = Path(os.environ.get("LOG_ROOT", "logs"))
    log_root.mkdir(parents=True, exist_ok=True)
    out = log_root / "ispring-probe.json"
    out.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(f"Wrote {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
