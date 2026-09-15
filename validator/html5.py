"""Local checks on a published iSpring HTML5 folder."""

from __future__ import annotations

import logging
from pathlib import Path

from utils.exceptions import OutputValidationError
from utils.paths import iter_files

logger = logging.getLogger(__name__)


def validate_ispring_output(output_dir: Path) -> list[dict]:
    if not output_dir.exists() or not output_dir.is_dir():
        raise OutputValidationError(
            f"publish output directory is missing: {output_dir}",
            user_message="iSpring did not produce an output folder.",
        )

    manifest = [
        {"path": relative, "size": absolute.stat().st_size}
        for absolute, relative in iter_files(output_dir)
    ]
    if not manifest:
        raise OutputValidationError(
            f"publish output directory is empty: {output_dir}",
            user_message="iSpring produced an empty output folder.",
        )

    index = output_dir / "index.html"
    if not index.exists() or index.stat().st_size == 0:
        raise OutputValidationError(
            "publish output is missing a non-empty index.html",
            user_message="iSpring output is missing index.html.",
        )

    has_asset_dir = any(
        "/" in row["path"] and (output_dir / row["path"].split("/")[0]).is_dir()
        for row in manifest
    )
    if not has_asset_dir:
        raise OutputValidationError(
            "publish output has no asset subdirectory; refusing a stub package",
            user_message="iSpring output looks incomplete.",
        )

    logger.info(
        "output validated",
        extra={"file_count": len(manifest), "stage": "output_validate"},
    )
    return manifest
