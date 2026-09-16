"""Third iSpring probe: dump the real window and control tree.

The add-in exposes no automation object and no command line, so the publish
has to be driven through the interface. This dumps what is actually on screen
so the adapter targets real control names instead of guesses.

Run it twice on the Windows machine, with PowerPoint already open:

    1. Open a deck, click the iSpring Suite 11 ribbon tab, then run:
           python scripts\\probe_ispring_ui.py --label ribbon
    2. Open the iSpring Publish window and stop on the HTML5 page, then run:
           python scripts\\probe_ispring_ui.py --label publish

Nothing is clicked and nothing is published. Each run writes
LOG_ROOT/ispring-ui-<label>.txt.
"""

from __future__ import annotations

import argparse
import io
import os
import sys
from contextlib import redirect_stdout
from datetime import datetime, timezone
from pathlib import Path

INTERESTING_PROCESSES = {
    "powerpnt.exe",
    "ispringlauncher.exe",
    "ispringsvr.exe",
    "ispringpreview.exe",
    "comcefview.exe",
    "comlauncher.exe",
    "infownd.exe",
}

INTERESTING_TITLE_WORDS = ("ispring", "publish", "powerpoint")

MAX_CHARS_PER_WINDOW = 400_000


def _process_name(pid: int) -> str:
    try:
        import psutil  # type: ignore

        return psutil.Process(pid).name()
    except Exception:  # noqa: BLE001
        return ""


def _is_interesting(process_name: str, title: str) -> bool:
    if process_name.lower() in INTERESTING_PROCESSES:
        return True
    lowered = (title or "").lower()
    return any(word in lowered for word in INTERESTING_TITLE_WORDS)


def _dump_window(window, depth: int) -> str:
    """Capture pywinauto's control identifiers for one top-level window."""
    buffer = io.StringIO()
    try:
        with redirect_stdout(buffer):
            window.print_control_identifiers(depth=depth)
    except Exception as exc:  # noqa: BLE001
        return f"    <could not read tree: {type(exc).__name__}: {exc}>\n"
    text = buffer.getvalue()
    if len(text) > MAX_CHARS_PER_WINDOW:
        text = text[:MAX_CHARS_PER_WINDOW] + "\n    <truncated>\n"
    return text


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Dump iSpring/PowerPoint UI trees.")
    parser.add_argument(
        "--label",
        default="ui",
        help="Name for this dump, e.g. ribbon or publish",
    )
    parser.add_argument(
        "--depth",
        type=int,
        default=6,
        help="How deep to walk each window (default 6; raise if a panel looks empty)",
    )
    parser.add_argument(
        "--all-windows",
        action="store_true",
        help="Dump every visible top-level window, not just iSpring/PowerPoint ones",
    )
    args = parser.parse_args(argv)

    if sys.platform != "win32":
        print("This probe needs the Windows machine.", file=sys.stderr)
        return 2

    try:
        from pywinauto import Desktop  # type: ignore
    except ImportError:
        print(
            "pywinauto is missing. Install it with: pip install pywinauto==0.6.8",
            file=sys.stderr,
        )
        return 2

    desktop = Desktop(backend="uia")
    sections: list[str] = [
        f"generated_at: {datetime.now(timezone.utc).isoformat()}",
        f"label: {args.label}",
        f"depth: {args.depth}",
        "",
    ]

    seen = 0
    for window in desktop.windows():
        try:
            title = window.window_text()
            pid = window.process_id()
        except Exception:  # noqa: BLE001
            continue
        process_name = _process_name(pid)
        if not args.all_windows and not _is_interesting(process_name, title):
            continue
        seen += 1
        try:
            rectangle = window.rectangle()
            class_name = window.class_name()
            control_type = window.element_info.control_type
            visible = window.is_visible()
        except Exception:  # noqa: BLE001
            rectangle = class_name = control_type = visible = "?"
        sections.append("=" * 78)
        sections.append(f"window: {title!r}")
        sections.append(f"  process: {process_name} (pid {pid})")
        sections.append(f"  class_name: {class_name}")
        sections.append(f"  control_type: {control_type}")
        sections.append(f"  visible: {visible}  rectangle: {rectangle}")
        sections.append("-" * 78)
        sections.append(_dump_window(window, args.depth))

    if seen == 0:
        sections.append(
            "No matching window found. Is PowerPoint open and on screen? "
            "A minimised window is fine, a locked screen is not."
        )

    log_root = Path(os.environ.get("LOG_ROOT", "logs"))
    log_root.mkdir(parents=True, exist_ok=True)
    out = log_root / f"ispring-ui-{args.label}.txt"
    out.write_text("\n".join(sections), encoding="utf-8", errors="replace")
    print(f"Wrote {out} ({seen} window(s))")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
