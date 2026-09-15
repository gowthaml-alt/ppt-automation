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
    def get_next_job(self):
        raise JobFetchError("down")


def _settings(**overrides):
    values = {
        "backend_base_url": "https://backend.example.com",
        "temp_root": "/tmp/jobs",
        "log_root": "/tmp/logs",
        "poll_interval_seconds": 1,
    }
    values.update(overrides)
    return Settings(**values)


def test_empty_queue_sleeps_and_does_not_process():
    sleeps = []
    backend = FakeBackend(jobs=[None])
    pipeline = RecordingPipeline()
    loop = WorkerLoop(_settings(), backend, pipeline, sleeper=sleeps.append)
    loop.run_once()
    assert pipeline.processed == []
    assert sleeps == [1]


def test_available_job_is_processed_once_without_sleeping():
    sleeps = []
    backend = FakeBackend(jobs=[make_job()])
    pipeline = RecordingPipeline()
    loop = WorkerLoop(_settings(), backend, pipeline, sleeper=sleeps.append)
    loop.run_once()
    assert pipeline.processed == [101]
    assert sleeps == []
    assert loop.busy is False


def test_second_job_is_not_fetched_until_first_finishes():
    backend = FakeBackend(jobs=[make_job(), make_job(queue_id=102)])
    pipeline = RecordingPipeline()
    loop = WorkerLoop(_settings(), backend, pipeline, sleeper=lambda _: None)
    loop.run_once()
    assert pipeline.processed == [101]
    loop.run_once()
    assert pipeline.processed == [101, 102]


def test_fetch_error_sleeps_without_processing():
    sleeps = []
    pipeline = RecordingPipeline()
    loop = WorkerLoop(
        _settings(), FetchErrorBackend(), pipeline, sleeper=sleeps.append
    )
    loop.run_once()
    assert pipeline.processed == []
    assert sleeps == [1]


def test_processing_failure_clears_busy_flag():
    backend = FakeBackend(jobs=[make_job()])
    pipeline = RecordingPipeline()
    pipeline.fail = True
    loop = WorkerLoop(_settings(), backend, pipeline, sleeper=lambda _: None)
    loop.run_once()
    assert loop.busy is False
