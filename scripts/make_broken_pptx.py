"""Make a damaged copy of a deck, for testing the repair path.

    python scripts\\make_broken_pptx.py --pptx "C:\\test\\sample.pptx" --mode slide-xml

Modes, from most to least repairable:

``slide-xml``  a valid zip whose first slide is malformed. This is the one
               PowerPoint offers to repair, so it is the useful test.
``drop-part``  a valid zip with a slide removed.
``truncate``   the file cut in half. Usually too broken to repair — good for
               checking that a hopeless file fails cleanly instead of hanging.

The original is never modified; a new file is written beside it.
"""

from __future__ import annotations

import argparse
import shutil
import sys
import zipfile
from pathlib import Path

BROKEN_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    b"<p:cSld><p:spTree><p:nvGrpSpPr>"  # deliberately never closed
)


def rewrite(source: Path, target: Path, *, damage_slide: bool, drop_slide: bool) -> None:
    """Copy the package, damaging or dropping the first slide on the way."""
    with zipfile.ZipFile(source) as original:
        names = original.namelist()
        slides = sorted(n for n in names if n.startswith("ppt/slides/slide"))
        if not slides:
            raise SystemExit("this file has no slides to damage")
        victim = slides[0]
        with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as broken:
            for name in names:
                if name == victim:
                    if drop_slide:
                        continue
                    if damage_slide:
                        broken.writestr(name, BROKEN_XML)
                        continue
                broken.writestr(name, original.read(name))
    print(f"damaged {victim} in {target.name}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Make a damaged copy of a .pptx")
    parser.add_argument("--pptx", required=True, help="A healthy .pptx to copy")
    parser.add_argument(
        "--mode",
        default="slide-xml",
        choices=("slide-xml", "drop-part", "truncate"),
        help="How to damage it (default: slide-xml, the repairable one)",
    )
    parser.add_argument("--out", default="", help="Where to write it")
    args = parser.parse_args(argv)

    source = Path(args.pptx).expanduser().resolve()
    if not source.is_file():
        print(f"not found: {source}", file=sys.stderr)
        return 1
    target = Path(args.out).expanduser().resolve() if args.out else source.with_name(
        f"{source.stem}-broken-{args.mode}.pptx"
    )

    if args.mode == "truncate":
        data = source.read_bytes()
        target.write_bytes(data[: len(data) // 2])
        print(f"cut {source.name} in half")
    else:
        shutil.copy2(source, target)
        rewrite(
            source,
            target,
            damage_slide=args.mode == "slide-xml",
            drop_slide=args.mode == "drop-part",
        )

    print(f"wrote {target}")
    print(f"size {target.stat().st_size} bytes")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
