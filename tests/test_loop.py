from api_client.jobs import JobFetchError
from config.settings import Settings
from tests.fakes import FakeBackend, make_job
from worker.loop import WorkerLoop


class RecordingPipeline:
    def __init__(self):
        self.processed = []
        self.fail = False

    def process(self, job):
        self.processed.append(job.queue_id)
        if self.fail:
            from utils.exceptions import DownloadError

            raise DownloadError("boom")


class FetchErrorBackend:
    def get_next_job(self, base_url=None):
        raise JobFetchError("down")


class PerServerBackend:
    """A queue per server. get_next_job pops from the one it is asked for."""

    def __init__(self, queues):
        self.queues = queues
        self.asked = []
        self.dead = set()

    def get_next_job(self, base_url=None):
        key = base_url or "shared"
        self.asked.append(key)
        if key in self.dead:
            raise JobFetchError("down")
        queue = self.queues.get(key, [])
        return queue.pop(0) if queue else None


def _settings(tmp_path, **overrides):
    values = {
        "backend_base_url": "https://api.example.com/nuSource/api/v1",
        "temp_root": str(tmp_path / "jobs"),
        "log_root": str(tmp_path / "logs"),
        "poll_interval_seconds": 1,
        "private_base_urls": [],
    }
    values.update(overrides)
    return Settings(**values)


def test_empty_queue_sleeps_and_does_not_process(tmp_path):
    sleeps = []
    backend = FakeBackend(jobs=[None])
    pipeline = RecordingPipeline()
    loop = WorkerLoop(_settings(tmp_path), backend, pipeline, sleeper=sleeps.append)
    loop.run_once()
    assert pipeline.processed == []
    assert sleeps == [1]


def test_available_job_is_processed_once_without_sleeping(tmp_path):
    sleeps = []
    backend = FakeBackend(jobs=[make_job()])
    pipeline = RecordingPipeline()
    loop = WorkerLoop(_settings(tmp_path), backend, pipeline, sleeper=sleeps.append)
    loop.run_once()
    assert pipeline.processed == [101]
    assert sleeps == []
    assert loop.busy is False


def test_second_job_is_not_fetched_until_first_finishes(tmp_path):
    backend = FakeBackend(jobs=[make_job(), make_job(queue_id=102)])
    pipeline = RecordingPipeline()
    loop = WorkerLoop(_settings(tmp_path), backend, pipeline, sleeper=lambda _: None)
    loop.run_once()
    assert pipeline.processed == [101]
    loop.run_once()
    assert pipeline.processed == [101, 102]


def test_fetch_error_sleeps_without_processing(tmp_path):
    sleeps = []
    pipeline = RecordingPipeline()
    loop = WorkerLoop(
        _settings(tmp_path), FetchErrorBackend(), pipeline, sleeper=sleeps.append
    )
    loop.run_once()
    assert pipeline.processed == []
    assert sleeps == [1]


def test_processing_failure_clears_busy_flag(tmp_path):
    backend = FakeBackend(jobs=[make_job()])
    pipeline = RecordingPipeline()
    pipeline.fail = True
    loop = WorkerLoop(_settings(tmp_path), backend, pipeline, sleeper=lambda _: None)
    loop.run_once()
    assert loop.busy is False


def test_private_server_is_emptied_before_the_next_one(tmp_path):
    backend = PerServerBackend(
        {
            "https://a.example.com": [make_job(queue_id=1), make_job(queue_id=2)],
            "https://b.example.com": [make_job(queue_id=3)],
        }
    )
    pipeline = RecordingPipeline()
    settings = _settings(
        tmp_path,
        private_base_urls=["https://a.example.com", "https://b.example.com"],
    )
    loop = WorkerLoop(settings, backend, pipeline, sleeper=lambda _: None)
    loop.run_once()
    assert pipeline.processed == [1, 2, 3]


def test_shared_queue_is_checked_between_private_rounds(tmp_path):
    ticks = iter([0, 0, 0, 0, 0, 0, 0, 0])
    backend = PerServerBackend(
        {
            "https://a.example.com": [make_job(queue_id=1)],
            "shared": [make_job(queue_id=9)],
        }
    )
    pipeline = RecordingPipeline()
    settings = _settings(tmp_path, private_base_urls=["https://a.example.com"])
    loop = WorkerLoop(
        settings,
        backend,
        pipeline,
        sleeper=lambda _: None,
        clock=lambda: next(ticks),
    )
    loop.run_once()
    assert pipeline.processed == [1]
    # The private clock is now 600s away, so the next round is the shared one.
    loop.run_once()
    assert pipeline.processed == [1, 9]


def test_dead_private_server_is_skipped_not_fatal(tmp_path):
    backend = PerServerBackend(
        {"https://b.example.com": [make_job(queue_id=5)]}
    )
    backend.dead.add("https://a.example.com")
    pipeline = RecordingPipeline()
    settings = _settings(
        tmp_path,
        private_base_urls=["https://a.example.com", "https://b.example.com"],
    )
    loop = WorkerLoop(settings, backend, pipeline, sleeper=lambda _: None)
    loop.run_once()
    assert pipeline.processed == [5]


def test_private_round_stops_at_the_cap(tmp_path):
    from worker.loop import PRIVATE_MAX_PER_ROUND

    backend = PerServerBackend(
        {
            "https://a.example.com": [
                make_job(queue_id=n) for n in range(PRIVATE_MAX_PER_ROUND + 5)
            ]
        }
    )
    pipeline = RecordingPipeline()
    settings = _settings(tmp_path, private_base_urls=["https://a.example.com"])
    loop = WorkerLoop(settings, backend, pipeline, sleeper=lambda _: None)
    loop.run_once()
    assert len(pipeline.processed) == PRIVATE_MAX_PER_ROUND
