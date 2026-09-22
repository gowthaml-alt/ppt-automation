"""iSpring publishing adapters.

Probing the installed Suite 11 (scripts/probe_ispring.py and
scripts/probe_ispring_api.py) established that iSpring exposes no automation
object, no add-in macro and no command line. Interface automation is therefore
the only way to publish, and it lives in ``publisher.ispring_cloud``.

Adapters:

``not_configured``  the default; fails loudly
``fake``            copies a checked-in fixture; unit tests only
``uia``             the real one: PowerPoint's interface plus iSpring Cloud
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

_NOT_CONFIGURED_HINT = (
    "iSpring publishing is not configured. Set ISPRING_ADAPTER=uia to publish "
    "through the PowerPoint interface to iSpring Cloud."
)


@dataclass(frozen=True)
class ISpringAvailability:
    installed: bool
    detail: str


class ISpringPublisher(Protocol):
    def is_available(self) -> ISpringAvailability: ...

    def publish(self, pptx: Path, output_dir: Path, timeout_s: int): ...


class NotConfiguredPublisher:
    def is_available(self) -> ISpringAvailability:
        return ISpringAvailability(installed=False, detail=_NOT_CONFIGURED_HINT)

    def publish(self, pptx: Path, output_dir: Path, timeout_s: int) -> Path:
        raise ISpringNotConfiguredError(_NOT_CONFIGURED_HINT)


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


class CloudPublisher:
    """Publishes to iSpring Cloud and reports the material's embed URL.

    Nothing is written to ``output_dir``: the content lives in iSpring Cloud,
    so this returns a URL where the other adapters return a folder.
    """

    def __init__(self, settings: Settings) -> None:
        self._settings = settings

    def is_available(self) -> ISpringAvailability:
        import sys

        if sys.platform != "win32":
            return ISpringAvailability(
                installed=False,
                detail="iSpring Cloud publishing only runs on the Windows worker",
            )
        install_path = self._settings.ispring_install_path
        if install_path and not Path(install_path).exists():
            return ISpringAvailability(
                installed=False, detail=f"ISPRING_INSTALL_PATH does not exist: {install_path}"
            )
        return ISpringAvailability(
            installed=True, detail="Windows host; iSpring Suite interface automation"
        )

    def publish_to_cloud(self, pptx: Path, *, institution: str, content_name: str):
        from publisher.ispring_cloud import publish_to_cloud

        return publish_to_cloud(
            Path(pptx),
            institution=institution,
            content_name=content_name,
            parent_folders=self._settings.ispring_parent_folders,
            cdp_url=self._settings.ispring_chrome_cdp_url,
            publish_timeout_s=self._settings.ispring_publish_timeout_seconds,
            browser_profile_dir=self._settings.ispring_chrome_profile_dir,
            browser_path=self._settings.ispring_chrome_path,
            cloud_url=self._settings.ispring_cloud_url,
        )

    def publish(self, pptx: Path, output_dir: Path, timeout_s: int):
        raise ISpringNotConfiguredError(
            "the uia adapter publishes to iSpring Cloud, not to a folder; call "
            "publish_to_cloud() with the institution and content name"
        )


def get_publisher(settings: Settings) -> ISpringPublisher:
    choice = settings.ispring_adapter
    if choice == "fake":
        return FakePublisher()
    if choice == "not_configured":
        return NotConfiguredPublisher()
    if choice == "uia":
        return CloudPublisher(settings)
    raise ISpringNotConfiguredError(
        f"unknown ISPRING_ADAPTER {choice!r}. {_NOT_CONFIGURED_HINT}"
    )
