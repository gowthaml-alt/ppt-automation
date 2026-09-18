"""Show, or clear, this worker's health.

    python scripts\\worker_health.py            what state it is in
    python scripts\\worker_health.py --clear    put it back in service

A worker stops asking for jobs after three failures in a row that point at
the machine — a missing iSpring add-in, a browser that is signed out. Fix
the machine first, then clear it here.
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from config.settings import get_settings  # noqa: E402
from worker import health  # noqa: E402


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--clear", action="store_true", help="Put the worker back in service")
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    root = get_settings().log_root
    if args.clear:
        health.clear(root)
        print("worker is back in service")
        return 0

    state = health.load(root)
    print(f"state file: {health.state_path(root)}")
    print(f"healthy:    {state.healthy}")
    print(f"failures:   {state.consecutive_failures} in a row")
    if state.reason:
        print(f"reason:     {state.reason}")
    if state.history:
        print("\nrecent failures")
        for line in state.history[-5:]:
            print(f"  {line}")
    return 0 if state.healthy else 1


if __name__ == "__main__":
    raise SystemExit(main())
