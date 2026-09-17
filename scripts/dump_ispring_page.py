"""Write out everything the iSpring Cloud page is showing right now.

The Windows half has scripts/probe_ispring_ui.py for dumping a window's
control tree. This is the same thing for the browser: attach to the Chrome
the automation drives, and write every visible element, every open popup and
the page text to a file that can be pasted back.

    python scripts\\dump_ispring_page.py
    python scripts\\dump_ispring_page.py --search "Kickoff and Advanced Prompting"
    python scripts\\dump_ispring_page.py --share "Kickoff and Advanced Prompting"

--search types the term into the library search box first; --share also opens
that row's three-dot menu and clicks Share, so the dump holds the share popup
with the embed code in it. Nothing is closed afterwards: the point is to leave
the screen exactly as the dump describes it.
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from config.settings import get_settings  # noqa: E402
from publisher.ispring_cloud import (  # noqa: E402
    debug_port_open,
    dump_page,
    locate_material,
    open_share,
    search_library,
    start_browser,
    visible_menu_labels,
)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--search", default="", help="type this into the library search")
    parser.add_argument("--share", default="", help="open this material's Share popup")
    parser.add_argument("--institution", default="", help="folder the material is in")
    parser.add_argument("--label", default="manual", help="name for the dump file")
    parser.add_argument("--cdp", default="", help="override the debug port URL")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(levelname)-7s %(message)s")
    settings = get_settings()
    cdp_url = args.cdp or settings.ispring_chrome_cdp_url

    from playwright.sync_api import sync_playwright

    if not debug_port_open(cdp_url):
        print(f"nothing on {cdp_url}; starting the share browser")
        start_browser(
            cdp_url, settings.ispring_chrome_profile_dir, settings.ispring_chrome_path
        )

    with sync_playwright() as playwright:
        try:
            browser = playwright.chromium.connect_over_cdp(cdp_url)
        except Exception as exc:  # noqa: BLE001
            print(f"could not attach to Chrome at {cdp_url}: {exc}")
            print("start it with scripts\\start_ispring_chrome.cmd and sign in once.")
            return 2
        if not browser.contexts:
            print("attached, but that Chrome has no window open")
            return 2
        context = browser.contexts[0]

        pages = [p for p in context.pages if "ispring" in (p.url or "").lower()]
        if pages:
            page = pages[-1]
        else:
            page = context.pages[-1] if context.pages else context.new_page()
            page.goto(settings.ispring_cloud_url, wait_until="domcontentloaded", timeout=60000)
        page.bring_to_front()
        page.wait_for_timeout(2000)

        material = args.share or args.search
        if args.share:
            if not locate_material(page, material, args.institution):
                print(f"{material!r} was not found in the library")
                dump_page(page, f"{args.label}-not-found")
                return 1
            if not open_share(page, material):
                print(f"no Share in the row menu; on screen: {visible_menu_labels(page)[:20]}")
                dump_page(page, f"{args.label}-no-share")
                return 1
            page.wait_for_timeout(2500)
        elif args.search:
            search_library(page, args.search)
            page.wait_for_timeout(1500)

        target = dump_page(page, args.label)
        if not target:
            print("the dump could not be written")
            return 1
        print(f"\nwrote {target}")
        print("paste that file back, and the screenshot beside it if it helps.")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
