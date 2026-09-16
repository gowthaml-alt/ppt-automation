"""Third iSpring probe: dump the real window and control tree.

The add-in exposes no automation object and no command line, so the publish
has to be driven through the interface. This dumps what is actually on screen
so the adapter targets real control names instead of guesses.

Run it on the Windows machine with PowerPoint already open and in front:

    1. Open a deck, click the iSpring Suite 11 ribbon tab, then run:
           python scripts\\probe_ispring_ui.py --label ribbon
    2. Open the iSpring Publish window and stop on the HTML5 page, then run:
           python scripts\\probe_ispring_ui.py --label publish

Nothing is clicked and nothing is published. Each run writes
LOG_ROOT/ispring-ui-<label>.txt, which always begins with a list of every
visible window it could see.
"""

from __future__ import annotations

import argparse
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

# Windows that are worth walking, by process name.
TARGET_PROCESSES = {
    "powerpnt.exe",
    "ispringlauncher.exe",
    "ispringsvr.exe",
    "ispringpreview.exe",
    "comcefview.exe",
    "comlauncher.exe",
    "infownd.exe",
}

# Never walk these, even if the title mentions iSpring: they are the terminal
# this probe is running in, or the editor the path was copied from.
IGNORED_PROCESSES = {
    "windowsterminal.exe",
    "cmd.exe",
    "conhost.exe",
    "powershell.exe",
    "pwsh.exe",
    "code.exe",
    "explorer.exe",
    "python.exe",
    "notepad.exe",
    "chrome.exe",
    "msedge.exe",
}

TITLE_WORDS = ("ispring", "publish", "powerpoint")

MAX_LINES_PER_WINDOW = 4000


def _process_name(pid: int) -> str:
    try:
        import psutil  # type: ignore

        return psutil.Process(pid).name()
    except Exception:  # noqa: BLE001
        return ""


def _text(value) -> str:
    try:
        return str(value)
    except Exception:  # noqa: BLE001
        return "?"


def _describe(element) -> str:
    name = _text(getattr(element, "name", ""))
    control_type = _text(getattr(element, "control_type", ""))
    class_name = _text(getattr(element, "class_name", ""))
    automation_id = _text(getattr(element, "automation_id", ""))
    try:
        rect = _text(element.rectangle)
    except Exception:  # noqa: BLE001
        rect = "?"
    parts = [f"{control_type} {name!r}"]
    if automation_id:
        parts.append(f"automation_id={automation_id!r}")
    if class_name:
        parts.append(f"class={class_name!r}")
    parts.append(f"rect={rect}")
    return "  ".join(parts)


def _walk(element, depth: int, max_depth: int, lines: list[str]) -> None:
    """Depth-first walk of the UIA tree. Version-proof: uses element_info only."""
    if len(lines) >= MAX_LINES_PER_WINDOW:
        return
    lines.append("    " + ("  " * depth) + _describe(element))
    if depth >= max_depth:
        return
    try:
        children = element.children()
    except Exception as exc:  # noqa: BLE001
        lines.append("    " + ("  " * (depth + 1)) + f"<children failed: {exc}>")
        return
    for child in children:
        _walk(child, depth + 1, max_depth, lines)
        if len(lines) >= MAX_LINES_PER_WINDOW:
            lines.append("    <truncated: raise --depth carefully or narrow --only>")
            return


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Dump iSpring/PowerPoint UI trees.")
    parser.add_argument("--label", default="ui", help="Name for this dump")
    parser.add_argument(
        "--depth",
        type=int,
        default=8,
        help="How deep to walk each window (default 8)",
    )
    parser.add_argument(
        "--only",
        default="",
        help="Only walk windows whose title contains this text (case-insensitive)",
    )
    parser.add_argument(
        "--all-windows",
        action="store_true",
        help="Walk every visible window, not just PowerPoint and iSpring ones",
    )
    parser.add_argument(
        "--wait",
        type=int,
        default=0,
        help=(
            "Seconds to wait before reading the screen. Use this to start the "
            "probe, then switch to PowerPoint and open the window you want dumped."
        ),
    )
    args = parser.parse_args(argv)

    if sys.platform != "win32":
        print("This probe needs the Windows machine.", file=sys.stderr)
        return 2

    try:
        from pywinauto.uia_element_info import UIAElementInfo  # type: ignore
    except ImportError:
        print(
            "pywinauto is missing. Install it with: pip install pywinauto==0.6.8",
            file=sys.stderr,
        )
        return 2

    if args.wait > 0:
        import time

        print(
            f"Waiting {args.wait}s. Switch to PowerPoint now and open the window "
            "you want dumped."
        )
        for remaining in range(args.wait, 0, -1):
            print(f"  {remaining:>3}s ", end="\r", flush=True)
            time.sleep(1)
        print("Reading the screen now.      ")

    own_pid = os.getpid()
    desktop = UIAElementInfo()  # the desktop root
    try:
        top_level = desktop.children()
    except Exception as exc:  # noqa: BLE001
        print(f"Could not read the desktop: {exc}", file=sys.stderr)
        return 1

    inventory: list[str] = []
    walked = 0
    sections: list[str] = []

    for element in top_level:
        title = _text(getattr(element, "name", ""))
        try:
            pid = int(element.process_id)
        except Exception:  # noqa: BLE001
            pid = -1
        process_name = _process_name(pid)
        inventory.append(f"  {process_name or '?':<24} pid={pid:<8} {title!r}")

        if pid == own_pid:
            continue
        lowered_process = process_name.lower()
        lowered_title = title.lower()

        if lowered_process in IGNORED_PROCESSES and not args.all_windows:
            # The terminal this probe runs in has the search word in its own
            # title, so it must never match.
            continue

        if args.only:
            wanted = args.only.lower() in lowered_title
        elif args.all_windows:
            wanted = True
        else:
            wanted = lowered_process in TARGET_PROCESSES or (
                lowered_process not in IGNORED_PROCESSES
                and any(word in lowered_title for word in TITLE_WORDS)
            )
        if not wanted:
            continue

        walked += 1
        lines: list[str] = []
        _walk(element, 0, args.depth, lines)
        sections.append("=" * 78)
        sections.append(f"window: {title!r}  ({process_name}, pid {pid})")
        sections.append("-" * 78)
        sections.extend(lines)
        sections.append("")

    header = [
        f"generated_at: {datetime.now(timezone.utc).isoformat()}",
        f"label: {args.label}   depth: {args.depth}   walked: {walked}",
        "",
        "ALL VISIBLE TOP-LEVEL WINDOWS",
        *inventory,
        "",
    ]
    if walked == 0:
        header.append(
            "Nothing matched. If PowerPoint is open, find it in the list above "
            "and re-run with --only \"<part of its title>\"."
        )
        header.append("")

    log_root = Path(os.environ.get("LOG_ROOT", "logs"))
    log_root.mkdir(parents=True, exist_ok=True)
    out = log_root / f"ispring-ui-{args.label}.txt"
    out.write_text("\n".join(header + sections), encoding="utf-8", errors="replace")
    print(f"Wrote {out} ({walked} window(s) walked, {len(inventory)} seen)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
