"""Check, and optionally repair, the iSpring PowerPoint add-in.

    python scripts\\fix_ispring_addin.py         show what is wrong
    python scripts\\fix_ispring_addin.py --fix   put it right

Every job does this check by itself now; this is for looking by hand, and for
the case where the repair needs an elevated prompt.

Close PowerPoint first: it rewrites these registry keys when it exits.
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from powerpoint.addin import check_addin, ensure_addin_enabled  # noqa: E402


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--fix", action="store_true", help="Apply the repairs")
    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    if sys.platform != "win32":
        print("This runs on the Windows machine.", file=sys.stderr)
        return 2

    state = check_addin()
    print("iSpring add-in registrations")
    for label in state.registered or ["  none found"]:
        print(f"  {label}")
    print("\nDemoted registrations")
    for label in state.demoted or ["  none"]:
        print(f"  {label}")
    print("\nOffice disabled items")
    for path, name in state.disabled_items or [("  none", "")]:
        print(f"  {path} {name}".rstrip())

    if state.healthy:
        print("\nNothing to fix.")
        return 0
    if not args.fix:
        print("\nClose PowerPoint, then run again with --fix")
        return 1
    return 0 if ensure_addin_enabled() else 1


if __name__ == "__main__":
    raise SystemExit(main())
