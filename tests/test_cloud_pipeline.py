from config.settings import Settings
from tests.fakes import make_job
from utils.exceptions import DownloadError, ISpringPublishingError
from worker.cloud_pipeline import CloudPipeline
from worker.health import should_take_work


class RecordingBackend:
    def __init__(self):
        self.results = []
        self.retry_calls = []

    def send_job_result(self, job, **kwargs):
        self.results.append({"job_id": job.job_id, **kwargs})


# What iSpring's share popup actually hands over.
EMBED_CODE = (
    '<iframe src="https://harshit.ispring.com/app/embed-player/abc" '
    'width="1280" height="720" frameborder="0" allowfullscreen></iframe>'
)


class FakeCloudJob:
    def __init__(self, fail=None, iframe_url="https://harshit.ispring.com/app/embed-player/abc"):
        self.fail = fail
        self.iframe_url = iframe_url
        self.calls = []

    def __call__(self, **kwargs):
        self.calls.append(kwargs)
        if self.fail is not None:
            raise self.fail
        return type(
            "R",
            (),
            {
                "iframe_url": self.iframe_url,
                "embed_code": EMBED_CODE,
                "content_name": str(kwargs.get("job_id")),
                "source_name": "source.pptx",
                "repaired": False,
                "elapsed_s": 1.0,
            },
        )()


def _settings(tmp_path, **overrides) -> Settings:
    values = {
        "backend_base_url": "https://api.example.com/nuSource/api/v1",
        "temp_root": str(tmp_path / "jobs"),
        "log_root": str(tmp_path / "logs"),
        "poll_interval_seconds": 1,
    }
    values.update(overrides)
    return Settings(**values)


def test_success_posts_status_2_with_the_whole_iframe(tmp_path):
    """The column holds the embed code, not the URL inside it.

    PreviewMaterial.tsx renders ispring_cloud_link as content, so a bare URL
    shows up as text on the page instead of the deck.
    """
    backend = RecordingBackend()
    cloud = FakeCloudJob()
    pipeline = CloudPipeline(_settings(tmp_path), backend, cloud_job=cloud)
    pipeline.process(make_job(job_id=123, material_name="Week 1 deck"))
    assert backend.results == [
        {
            "job_id": 123,
            "status": 2,
            "ispringcloud_link": EMBED_CODE,
        }
    ]
    assert not backend.retry_calls
    assert cloud.calls[0]["job_id"] == 123
    assert cloud.calls[0]["content_name"] == ""
    assert cloud.calls[0]["material_name"] == "Week 1 deck"


def test_failure_posts_status_3_and_does_not_retry(tmp_path):
    backend = RecordingBackend()
    cloud = FakeCloudJob(fail=DownloadError("empty", user_message="The downloaded file is empty."))
    pipeline = CloudPipeline(_settings(tmp_path), backend, cloud_job=cloud)
    pipeline.process(make_job(job_id=123))
    assert len(backend.results) == 1
    result = backend.results[0]
    assert result["job_id"] == 123
    assert result["status"] == 3
    assert result["stage"] == "download"
    assert result["error_code"] == "DOWNLOAD_FAILED"
    assert result["error_message"] == "The downloaded file is empty."
    assert not backend.retry_calls


def test_ispring_failure_maps_to_ispring_stage(tmp_path):
    backend = RecordingBackend()
    cloud = FakeCloudJob(fail=ISpringPublishingError("no ribbon", user_message="publish failed"))
    pipeline = CloudPipeline(_settings(tmp_path), backend, cloud_job=cloud)
    pipeline.process(make_job(job_id=123))
    assert backend.results[0]["stage"] == "ispring"
    assert backend.results[0]["error_code"] == "ISPRING_FAILED"


def test_publish_uses_job_id_not_material_name(tmp_path):
    backend = RecordingBackend()
    cloud = FakeCloudJob()
    pipeline = CloudPipeline(_settings(tmp_path), backend, cloud_job=cloud)
    pipeline.process(
        make_job(job_id=123, material_id=789, material_name="Week 1 deck")
    )
    assert str(cloud.calls[0]["job_id"]) == "123"
    assert cloud.calls[0]["material_name"] == "Week 1 deck"


def test_unhealthy_worker_does_not_claim_work(tmp_path, monkeypatch):
    from worker.health import record_failure
    from worker.loop import WorkerLoop

    for _ in range(3):
        record_failure(tmp_path / "logs", stage="ispring_publish", message="broken")
    assert should_take_work(tmp_path / "logs") is False

    class ClaimBackend:
        def __init__(self):
            self.calls = 0

        def get_next_job(self):
            self.calls += 1
            return make_job()

    backend = ClaimBackend()
    pipeline = CloudPipeline(
        _settings(tmp_path), RecordingBackend(), cloud_job=FakeCloudJob()
    )
    sleeps = []
    loop = WorkerLoop(_settings(tmp_path), backend, pipeline, sleeper=sleeps.append)
    loop.run_once()
    assert backend.calls == 0
    assert sleeps == [1]
