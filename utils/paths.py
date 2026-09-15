"""Filesystem helpers. Every job path derives from the integer queue_id."""

from __future__ import annotations

import shutil
from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path

SOURCE_FILENAME = "source.pptx"
PARTIAL_SUFFIX = ".part"


@dataclass(frozen=True)
class JobPaths:
    root: Path
    input_dir: Path
    working_dir: Path
    output_dir: Path
    log_dir: Path

    @property
    def source_pptx(self) -> Path:
        return self.input_dir / SOURCE_FILENAME

    @property
    def partial_pptx(self) -> Path:
        return self.input_dir / (SOURCE_FILENAME + PARTIAL_SUFFIX)


def build_job_paths(temp_root: str | Path, queue_id: int) -> JobPaths:
    """Build the directory layout for a job.

    ``queue_id`` must be an ``int``. Requiring an integer is what makes path
    traversal impossible: there is no string for a caller to smuggle
    separators through.
    """
    if isinstance(queue_id, bool) or not isinstance(queue_id, int):
        raise TypeError(f"queue_id must be an int, got {type(queue_id).__name__}")
    if queue_id <= 0:
        raise ValueError(f"queue_id must be positive, got {queue_id}")

    root = Path(temp_root) / str(queue_id)
    return JobPaths(
        root=root,
        input_dir=root / "input",
        working_dir=root / "working",
        output_dir=root / "output",
        log_dir=root / "logs",
    )


def create_job_dirs(paths: JobPaths) -> None:
    for directory in (
        paths.root,
        paths.input_dir,
        paths.working_dir,
        paths.output_dir,
        paths.log_dir,
    ):
        directory.mkdir(parents=True, exist_ok=True)


def assert_within(base: Path, candidate: Path) -> Path:
    """Return ``candidate`` resolved, or raise if it escapes ``base``."""
    base_resolved = Path(base).resolve()
    candidate_resolved = Path(candidate).resolve()
    if (
        base_resolved != candidate_resolved
        and base_resolved not in candidate_resolved.parents
    ):
        raise ValueError(f"path {candidate_resolved} is outside {base_resolved}")
    return candidate_resolved


def free_disk_gb(path: str | Path) -> float:
    """Free space in GiB on the volume containing ``path``."""
    target = Path(path)
    while not target.exists() and target != target.parent:
        target = target.parent
    return shutil.disk_usage(target).free / (1024**3)


def iter_files(root: Path) -> Iterator[tuple[Path, str]]:
    """Yield ``(absolute_path, posix_relative_path)`` for every file under root."""
    root = Path(root)
    for entry in sorted(root.rglob("*")):
        if entry.is_file():
            yield entry, entry.relative_to(root).as_posix()
