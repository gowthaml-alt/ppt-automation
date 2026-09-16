"""Test doubles shared across the suite."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from api_client.jobs import Job, UploadResult
from utils.exceptions import (
    CallbackError,
    DownloadError,
    ISpringPublishingError,
    OutputUploadError,
)


def make_job(**overrides) -> Job:
    values = {
        "queue_id": 101,
        "material_id": 5001,
        "ppt_file_url": "https://storage.example.com/input/biology.pptx",
        "material_name": "Introduction to Biology",
        "institution_name": "ABC College",
    }
    values.update(overrides)
    return Job(**values)


@dataclass
class FakeBackend:
    jobs: list[Job | None] = field(default_factory=list)
    results: list[dict] = field(default_factory=list)
    fail_callback: bool = False

    def get_next_job(self) -> Job | None:
        if not self.jobs:
            return None
        return self.jobs.pop(0)

    def send_job_result(self, job: Job, *, status: str, iframe_url=None, error_message=None):
        if self.fail_callback:
            raise CallbackError("callback down")
        self.results.append(
            {
                "queue_id": job.queue_id,
                "material_id": job.material_id,
                "status": status,
                "iframe_url": iframe_url,
                "error_message": error_message,
            }
        )


@dataclass
class FakePowerPoint:
    started: bool = False
    opened: list[str] = field(default_factory=list)
    closed: int = 0
    quit_calls: int = 0
    fail_on: str | None = None
    last_timeout_s: float | None = None
    current_pptx: Path | None = None
    repaired_name: str | None = None

    def start(self) -> None:
        from utils.exceptions import PowerPointAutomationError

        if self.fail_on == "start":
            raise PowerPointAutomationError("fake start failure")
        self.started = True

    def open(self, path, timeout_s: float | None = None) -> None:
        from utils.exceptions import PowerPointAutomationError

        if self.fail_on == "open":
            raise PowerPointAutomationError("fake open failure")
        self.opened.append(str(path))
        self.last_timeout_s = timeout_s
        opened = Path(path)
        if self.repaired_name:
            opened = opened.with_name(self.repaired_name)
        self.current_pptx = opened

    def close(self) -> None:
        self.closed += 1

    def quit(self) -> None:
        self.quit_calls += 1
        self.current_pptx = None


@dataclass
class FakePublisher:
    fail: bool = False
    published: list[str] = field(default_factory=list)

    def publish(self, pptx, output_dir, timeout_s: int):
        if self.fail:
            raise ISpringPublishingError("fake publish failure")
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        (output_dir / "data").mkdir(exist_ok=True)
        (output_dir / "index.html").write_text("<html>ok</html>", encoding="utf-8")
        (output_dir / "data" / "app.js").write_text("ok", encoding="utf-8")
        self.published.append(str(pptx))
        return output_dir


@dataclass
class FakeStorage:
    fail: bool = False
    uploaded: list[str] = field(default_factory=list)

    def upload_directory(self, local, prefix: str) -> UploadResult:
        if self.fail:
            raise OutputUploadError("fake upload failure")
        self.uploaded.append(prefix)
        return UploadResult(
            output_path=prefix,
            entry_file="index.html",
            iframe_url=f"https://cdn.example.com/{prefix}/index.html",
        )


@dataclass
class FakeBrowser:
    fail: bool = False
    checked: list[str] = field(default_factory=list)

    def check(self, iframe_url: str) -> None:
        from utils.exceptions import BrowserTestError

        if self.fail:
            raise BrowserTestError("fake browser failure")
        self.checked.append(iframe_url)


@dataclass
class FakeSlack:
    calls: list[dict] = field(default_factory=list)

    def notify_failure(self, job: Job, *, stage: str, error_message: str) -> None:
        self.calls.append(
            {"queue_id": job.queue_id, "stage": stage, "error_message": error_message}
        )


def fake_downloader(url, paths, settings, **kwargs):
    if url.endswith("missing.pptx"):
        raise DownloadError("not found")
    paths.source_pptx.parent.mkdir(parents=True, exist_ok=True)
    paths.source_pptx.write_bytes(b"PK\x03\x04fake")
    return paths.source_pptx
