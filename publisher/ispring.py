"""iSpring publishing adapters.

The real publishing mechanism is unknown until scripts/probe_ispring.py is
run on the installed Cloud PC version. This module ships a loud default and
three empty adapters. Do not add guessed ProgIDs, executables, or ribbon
identifiers here.
"""

from __future__ import annotations

import logging
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Protocol

from config.settings import Settings
from utils.exceptions import ISpringNotConfiguredError

logger = logging.getLogger(__name__)

FIXTURE_DIR = (
    Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "ispring_output"
)

_PROBE_HINT = (
    "iSpring publishing is not configured. Run scripts/probe_ispring.py on the "
    "Windows Cloud PC, then set ISPRING_ADAPTER to the adapter named in that "
    "report and implement only the verified interface."
)


@dataclass(frozen=True)
class ISpringAvailability:
    installed: bool
    detail: str


class ISpringPublisher(Protocol):
    def is_available(self) -> ISpringAvailability: ...

    def publish(self, pptx: Path, output_dir: Path, timeout_s: int) -> Path: ...


class NotConfiguredPublisher:
    def is_available(self) -> ISpringAvailability:
        return ISpringAvailability(installed=False, detail=_PROBE_HINT)

    def publish(self, pptx: Path, output_dir: Path, timeout_s: int) -> Path:
        raise ISpringNotConfiguredError(_PROBE_HINT)


class UnverifiedAdapter(NotConfiguredPublisher):
    def __init__(self, name: str) -> None:
        self._name = name

    def is_available(self) -> ISpringAvailability:
        return ISpringAvailability(
            installed=False,
            detail=(
                f"ISPRING_ADAPTER={self._name} has no verified implementation. "
                f"{_PROBE_HINT}"
            ),
        )

    def publish(self, pptx: Path, output_dir: Path, timeout_s: int) -> Path:
        raise ISpringNotConfiguredError(
            f"ISPRING_ADAPTER={self._name} is a stub. {_PROBE_HINT}"
        )


class FakePublisher:
    """Copies the checked-in fixture package. Used by unit tests only."""

    def is_available(self) -> ISpringAvailability:
        return ISpringAvailability(installed=True, detail="fake publisher")

    def publish(self, pptx: Path, output_dir: Path, timeout_s: int) -> Path:
        if not FIXTURE_DIR.exists():
            raise ISpringNotConfiguredError(
                f"fake publisher fixture is missing: {FIXTURE_DIR}"
            )
        if output_dir.exists():
            shutil.rmtree(output_dir)
        shutil.copytree(FIXTURE_DIR, output_dir)
        logger.info(
            "fake iSpring publish copied fixture", extra={"stage": "ispring_publish"}
        )
        return output_dir


def get_publisher(settings: Settings) -> ISpringPublisher:
    choice = settings.ispring_adapter
    if choice == "fake":
        return FakePublisher()
    if choice == "not_configured":
        return NotConfiguredPublisher()
    if choice in {"vba", "uia", "cli"}:
        return UnverifiedAdapter(choice)
    raise ISpringNotConfiguredError(
        f"unknown ISPRING_ADAPTER {choice!r}. {_PROBE_HINT}"
    )
