"""Publish one material: a PPT URL in, an iSpring Cloud iframe URL out.

    python scripts\\publish_material.py ^
        --url "https://dme2wmiz2suov.cloudfront.net/User(1)/...pptx?Expires=..." ^
        --material-name "Kickoff and Advanced Prompting" ^
        --material-id 20242897 ^
        --institution-name "Demoacademy"

Downloads the file, opens it in PowerPoint (repairing it if it is damaged),
publishes it into the institution's project in iSpring Cloud, reads the share
iframe, and closes everything it opened.

Quote the URL. A signed URL contains & and ? and the shell will cut it in half
otherwise.

Before the first run, start the browser it talks to and sign in once:

    scripts\\start_ispring_chrome.cmd
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from config.settings import get_settings  # noqa: E402
from utils.exceptions import PptAutomationError  # noqa: E402
from utils.logging_config import bind_job_context, configure_logging  # noqa: E402
from worker import health  # noqa: E402
from worker.cloud_job import run_cloud_job  # noqa: E402


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Publish one material to iSpring Cloud and print its iframe URL."
    )
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--url", help="Signed URL of the .pptx (quote it)")
    source.add_argument("--pptx", help="Local .pptx, instead of downloading one")
    parser.add_argument("--material-name", required=True)
    parser.add_argument("--material-id", required=True)
    parser.add_argument(
        "--job-id",
        default="",
        help="Queue row id; the deck is published under it",
    )
    parser.add_argument("--institution-name", required=True)
    parser.add_argument(
        "--content-name",
        default="",
        help="Name in iSpring Cloud (default: the material name)",
    )
    parser.add_argument(
        "--keep-files",
        action="store_true",
        help="Leave the downloaded and repaired files on disk",
    )
    parser.add_argument(
        "--ignore-health",
        action="store_true",
        help="Run even when this worker is marked unhealthy",
    )
    args = parser.parse_args(argv)

    if sys.platform != "win32":
        print("This runs on the Windows worker.", file=sys.stderr)
        return 2

    settings = get_settings()
    configure_logging(settings)
    try:
        bind_job_context(queue_id=0, material_id=int(args.material_id))
    except (TypeError, ValueError):
        pass

    if not args.ignore_health and not health.should_take_work(settings.log_root):
        state = health.load(settings.log_root)
        print(
            f"\nSTOPPED: this worker is marked unhealthy ({state.reason}).\n"
            f"Fix the machine, then clear it:\n"
            f"    python scripts\\worker_health.py --clear",
            file=sys.stderr,
        )
        return 3

    try:
        result = run_cloud_job(
            url=args.url or "",
            pptx=args.pptx or "",
            material_name=args.material_name,
            material_id=args.material_id,
            institution_name=args.institution_name,
            settings=settings,
            job_id=args.job_id,
            content_name=args.content_name,
            keep_files=args.keep_files,
        )
    except PptAutomationError as exc:
        health.record_failure(
            settings.log_root, stage=exc.stage or "", message=exc.message or str(exc)
        )
        print(f"\nFAILED: {exc.user_message or exc}", file=sys.stderr)
        # The polite message alone is not enough to fix anything.
        if exc.message and exc.message != exc.user_message:
            print(f"detail: {exc.message}", file=sys.stderr)
        print(f"stage: {exc.stage}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("interrupted", file=sys.stderr)
        return 130
    except Exception as exc:  # noqa: BLE001
        health.record_failure(
            settings.log_root, stage="unknown", message=f"{type(exc).__name__}: {exc}"
        )
        print(f"\nFAILED (unexpected): {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1

    health.record_success(settings.log_root)
    print("\nPUBLISH OK")
    print(f"material_id={args.material_id}")
    print(f"published_as={result.content_name}")
    print(f"source_name={result.source_name}")
    print(f"repaired={result.repaired}")
    print(f"iframe_url={result.iframe_url}")
    print(f"embed_code={result.embed_code}")
    print(f"seconds={result.elapsed_s:.0f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
