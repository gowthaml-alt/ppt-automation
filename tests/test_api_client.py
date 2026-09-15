import json

import httpx
import pytest

from api_client.jobs import BackendClient, parse_job_payload, record_orphaned_output
from api_client.jobs import UploadResult
from config.settings import Settings
from tests.fakes import make_job
from utils.exceptions import CallbackError, JobFetchError


def _settings(tmp_path, **overrides) -> Settings:
    values = {
        "app_env": "development",
        "backend_base_url": "https://backend.example.com",
        "backend_get_job_path": "/api/ppt-jobs/next",
        "backend_callback_path": "/api/ppt-jobs/callback",
        "backend_request_timeout_seconds": 5,
        "get_job_retry_attempts": 1,
        "callback_retry_attempts": 1,
        "temp_root": str(tmp_path / "jobs"),
        "log_root": str(tmp_path / "logs"),
    }
    values.update(overrides)
    return Settings(**values)


SAMPLE = {
    "queue_id": 101,
    "material_id": 5001,
    "material_name": "Introduction to Biology",
    "institution_name": "ABC College",
    "ppt_file_url": "https://s3.example.com/in/a.pptx?X-Amz-Signature=secret",
}


def test_parse_job_payload_maps_required_fields():
    job = parse_job_payload(SAMPLE)
    assert job.queue_id == 101
    assert job.material_id == 5001
    assert "X-Amz-Signature" in job.ppt_file_url


def test_parse_job_payload_accepts_id_alias():
    job = parse_job_payload(
        {
            "id": 9,
            "material_id": 1,
            "ppt_file_url": "https://example.com/a.pptx",
        }
    )
    assert job.queue_id == 9


def test_parse_job_payload_unwraps_job_key():
    job = parse_job_payload({"job": SAMPLE})
    assert job.queue_id == 101


def test_parse_job_payload_empty_shapes():
    assert parse_job_payload(None) is None
    assert parse_job_payload({}) is None
    assert parse_job_payload({"job": None}) is None
    assert parse_job_payload([]) is None


def test_parse_job_payload_incomplete_raises():
    with pytest.raises(JobFetchError):
        parse_job_payload({"queue_id": 1, "material_id": 2})


def test_get_next_job_returns_job(tmp_path):
    def handler(request: httpx.Request) -> httpx.Response:
        assert request.method == "GET"
        assert str(request.url) == "https://backend.example.com/api/ppt-jobs/next"
        return httpx.Response(200, json=SAMPLE)

    client = BackendClient(
        _settings(tmp_path),
        client=httpx.Client(transport=httpx.MockTransport(handler)),
    )
    job = client.get_next_job()
    assert job.queue_id == 101


def test_get_next_job_204_is_empty(tmp_path):
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(204)

    client = BackendClient(
        _settings(tmp_path),
        client=httpx.Client(transport=httpx.MockTransport(handler)),
    )
    assert client.get_next_job() is None


def test_get_next_job_sends_optional_api_key(tmp_path):
    seen = {}

    def handler(request: httpx.Request) -> httpx.Response:
        seen["header"] = request.headers.get("x-api-key")
        return httpx.Response(204)

    client = BackendClient(
        _settings(tmp_path, backend_api_key="secret"),
        client=httpx.Client(transport=httpx.MockTransport(handler)),
    )
    client.get_next_job()
    assert seen["header"] == "secret"


def test_get_next_job_does_not_send_api_key_when_unset(tmp_path):
    seen = {}

    def handler(request: httpx.Request) -> httpx.Response:
        seen["header"] = request.headers.get("x-api-key")
        return httpx.Response(204)

    client = BackendClient(
        _settings(tmp_path),
        client=httpx.Client(transport=httpx.MockTransport(handler)),
    )
    client.get_next_job()
    assert seen["header"] is None


def test_callback_success_body(tmp_path):
    seen = {}

    def handler(request: httpx.Request) -> httpx.Response:
        seen["url"] = str(request.url)
        seen["body"] = json.loads(request.content)
        return httpx.Response(200, json={"ok": True})

    client = BackendClient(
        _settings(tmp_path),
        client=httpx.Client(transport=httpx.MockTransport(handler)),
    )
    client.send_job_result(
        make_job(),
        status="completed",
        iframe_url="https://cdn.example.com/ppt/5001/index.html",
    )
    assert seen["url"] == "https://backend.example.com/api/ppt-jobs/callback"
    assert seen["body"] == {
        "queue_id": 101,
        "material_id": 5001,
        "status": "completed",
        "iframe_url": "https://cdn.example.com/ppt/5001/index.html",
    }


def test_callback_failure_body(tmp_path):
    seen = {}

    def handler(request: httpx.Request) -> httpx.Response:
        seen["body"] = json.loads(request.content)
        return httpx.Response(200)

    client = BackendClient(
        _settings(tmp_path),
        client=httpx.Client(transport=httpx.MockTransport(handler)),
    )
    client.send_job_result(make_job(), status="failed", error_message="iSpring publishing failed")
    assert seen["body"]["status"] == "failed"
    assert seen["body"]["error_message"] == "iSpring publishing failed"
    assert "iframe_url" not in seen["body"]


def test_callback_non_2xx_raises(tmp_path):
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(500, text="nope")

    client = BackendClient(
        _settings(tmp_path),
        client=httpx.Client(transport=httpx.MockTransport(handler)),
    )
    with pytest.raises(CallbackError):
        client.send_job_result(make_job(), status="failed", error_message="x")


def test_orphaned_output_is_appended(tmp_path):
    log_root = tmp_path / "logs"
    path = record_orphaned_output(
        log_root,
        make_job(),
        UploadResult("ppt/5001", "index.html", "https://cdn.example.com/ppt/5001/index.html"),
    )
    row = json.loads(path.read_text(encoding="utf-8").splitlines()[0])
    assert row["queue_id"] == 101
    assert row["iframe_url"].endswith("index.html")
