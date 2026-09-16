"""One-shot local Windows test: pass a PPTX, upload HTML5, verify it loads.

Does not call the Node GET job API. Example:

    python scripts/process_local.py --pptx "C:\\test\\sample.pptx"
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from config.settings import get_settings  # noqa: E402
from utils.logging_config import configure_logging  # noqa: E402
from worker.local_run import run_local_job  # noqa: E402


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Process one local PPTX: copy, (optional) PowerPoint, "
        "publish, upload, then HTTP/Playwright-check the output."
    )
    parser.add_argument("--pptx", required=True, help="Path to a local .pptx file")
    parser.add_argument("--queue-id", type=int, default=1)
    parser.add_argument("--material-id", type=int, default=1)
    parser.add_argument("--material-name", default="local-test")
    parser.add_argument("--institution-name", default="local")
    parser.add_argument(
        "--skip-powerpoint",
        action="store_true",
        help="Do not launch PowerPoint (use on machines without it)",
    )
    parser.add_argument(
        "--real-ispring",
        action="store_true",
        help="Use ISPRING_ADAPTER from .env instead of the fake fixture publisher",
    )
    parser.add_argument(
        "--callback",
        action="store_true",
        help="Also POST the Node callback (off by default for laptop tests)",
    )
    parser.add_argument(
        "--playwright",
        action="store_true",
        help="After HTTP check, also open the page in Chromium",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    settings = get_settings()
    configure_logging(settings)

    result = run_local_job(
        Path(args.pptx),
        settings=settings,
        queue_id=args.queue_id,
        material_id=args.material_id,
        material_name=args.material_name,
        institution_name=args.institution_name,
        skip_powerpoint=args.skip_powerpoint,
        use_fake_ispring=not args.real_ispring,
        skip_callback=not args.callback,
        use_playwright=args.playwright,
    )
    if result.ok:
        print("LOCAL RUN PASSED")
        if result.local_index:
            print(f"local_index={result.local_index}")
        if result.iframe_url:
            print(f"iframe_url={result.iframe_url}")
        if result.checked_url:
            print(f"checked_url={result.checked_url}")
        print(f"browser_ok={result.browser_ok}")
        if not args.real_ispring:
            print(
                "NOTE: fake iSpring was used. HTML5 is the fixture package, "
                "not a conversion of your slides."
            )
        return 0
    print("LOCAL RUN FAILED")
    print(result.error or "unknown error")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
