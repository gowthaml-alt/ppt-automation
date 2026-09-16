"""Copy a local PPTX into the job directory. Used by the one-shot Windows test."""

from __future__ import annotations

import logging
import shutil
from pathlib import Path

from utils.exceptions import DownloadError
from utils.paths import JobPaths
from utils.validators import assert_looks_like_pptx, assert_valid_pptx_package

logger = logging.getLogger(__name__)


def copy_local_pptx(source: Path, paths: JobPaths) -> Path:
    """Copy ``source`` to ``input/source.pptx`` after validating it looks like a PPTX."""
    src = Path(source).expanduser()
    if not src.is_file():
        raise DownloadError(
            f"local PPTX does not exist: {src}",
            user_message="The local PowerPoint file was not found.",
        )
    dest = paths.source_pptx
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dest)
    try:
        assert_looks_like_pptx(dest)
        assert_valid_pptx_package(dest)
    except DownloadError:
        dest.unlink(missing_ok=True)
        raise
    logger.info(
        "copied local PPTX",
        extra={"bytes": dest.stat().st_size, "stage": "download", "source_name": src.name},
    )
    return dest
