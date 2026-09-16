"""Shrink an ispring-ui-*.txt dump down to just the dialog windows.

The full dump is mostly PowerPoint's ribbon, which we already know. This keeps
only the child dialogs (class '#32770') from each snapshot — the publish form,
the progress window, and whatever appears at the end.

    python scripts\\summarize_ui_log.py logs\\ispring-ui-cloud.txt

Writes <input>-short.txt next to the original and prints it.
"""

from __future__ import annotations

import sys
from pathlib import Path

DIALOG_MARKER = "class='#32770'"


def indent_of(line: str) -> int:
    return len(line) - len(line.lstrip(" "))


def summarize(text: str) -> str:
    out: list[str] = []
    lines = text.splitlines()
    index = 0
    snapshot_header = ""
    printed_header = True

    while index < len(lines):
        line = lines[index]
        stripped = line.strip()

        if stripped.startswith("# SNAPSHOT") or stripped.startswith("# generated_at"):
            snapshot_header = (
                snapshot_header + "\n" + stripped if snapshot_header else stripped
            )
            printed_header = False
            index += 1
            continue

        if DIALOG_MARKER in line:
            if not printed_header:
                out.append("")
                out.append("=" * 70)
                out.append(snapshot_header)
                out.append("=" * 70)
                snapshot_header = ""
                printed_header = True
            base = indent_of(line)
            out.append(line.rstrip())
            index += 1
            while index < len(lines):
                child = lines[index]
                if not child.strip():
                    index += 1
                    continue
                if indent_of(child) <= base:
                    break
                out.append(child.rstrip())
                index += 1
            continue

        index += 1

    if not out:
        return "No dialog windows found in this dump."
    return "\n".join(out)


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        print("usage: python scripts\\summarize_ui_log.py <log file>", file=sys.stderr)
        return 2
    source = Path(argv[1])
    if not source.is_file():
        print(f"not found: {source}", file=sys.stderr)
        return 2
    summary = summarize(source.read_text(encoding="utf-8", errors="replace"))
    target = source.with_name(source.stem + "-short.txt")
    target.write_text(summary, encoding="utf-8")
    print(summary)
    print(f"\n[written to {target}]")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
