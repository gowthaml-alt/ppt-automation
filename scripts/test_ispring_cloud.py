"""Read-only iSpring Cloud browser test for macOS Chrome.

Opens the official iSpring Cloud website with Playwright. It never talks to
MySQL, never starts the poller, and never automates PowerPoint.

    python scripts/test_ispring_cloud.py

Log in manually in the Chrome window if asked. The script does not upload,
delete, publish, or edit any cloud content.
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from browser_test.ispring_cloud import (  # noqa: E402
    DEFAULT_MATERIAL_NAME,
    DEFAULT_SEARCH_QUERY,
    ArtifactRun,
    ISpringCloudProbeError,
    login_credentials,
    run_cloud_probe,
)

DEFAULT_ARTIFACTS = ROOT / "test-artifacts"
DEFAULT_PROFILE = DEFAULT_ARTIFACTS / "chrome-profile"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Open Google Chrome and walk through iSpring Cloud: search "
            f"{DEFAULT_SEARCH_QUERY!r}, open {DEFAULT_MATERIAL_NAME!r}, "
            "and capture screenshots. Read-only."
        )
    )
    parser.add_argument(
        "--url",
        default=os.environ.get("ISPRING_CLOUD_URL", ""),
        help="Optional start URL. Default: Harshit's Team iSpring app",
    )
    parser.add_argument(
        "--search",
        default=DEFAULT_SEARCH_QUERY,
        help=f"Library search text (default: {DEFAULT_SEARCH_QUERY!r})",
    )
    parser.add_argument(
        "--material",
        default=DEFAULT_MATERIAL_NAME,
        help=f"Material name to open (default: {DEFAULT_MATERIAL_NAME!r})",
    )
    parser.add_argument(
        "--artifacts-dir",
        default=os.environ.get("ISPRING_CLOUD_ARTIFACTS_DIR", str(DEFAULT_ARTIFACTS)),
        help="Directory for screenshots, traces, and logs",
    )
    parser.add_argument(
        "--profile-dir",
        default=os.environ.get("ISPRING_CLOUD_PROFILE_DIR", str(DEFAULT_PROFILE)),
        help="Persistent Chrome profile so a manual login can be reused",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    artifacts = ArtifactRun.create(Path(args.artifacts_dir))
    supplied = login_credentials(os.environ)
    email = supplied[0] if supplied else None
    password = supplied[1] if supplied else None
    try:
        info = run_cloud_probe(
            artifacts=artifacts,
            profile_dir=Path(args.profile_dir),
            start_url=args.url or None,
            search_query=args.search,
            material_name=args.material,
            headless=False,
            pause=None if supplied else input,
            email=email,
            password=password,
        )
    except ISpringCloudProbeError as exc:
        print(f"ISPRING CLOUD TEST FAILED: {exc}", file=sys.stderr)
        print(f"artifacts={artifacts.root}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("interrupted", file=sys.stderr)
        return 130

    print("ISPRING CLOUD TEST PASSED")
    print(f"title={info.title}")
    print(f"url={info.url}")
    print(f"material_id={info.material_id or ''}")
    print(f"share_controls={', '.join(info.share_controls)}")
    print(f"artifacts={artifacts.root}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
