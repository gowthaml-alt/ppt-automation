import json
from urllib.parse import parse_qs

import httpx
import pytest

from api_client.jobs import (
    NEXT_PATH,
    RESULT_PATH,
    RETRY_PATH,
    TOKEN_HEADER,
    BackendClient,
    error_code_for,
    parse_next_response,
    record_orphaned_output,
    result_stage,
)
from api_client.jobs import UploadResult
from config.settings import Settings
from tests.fakes import make_job
from utils.exceptions import (
    CallbackError,
    DownloadError,
    ISpringPublishingError,
    JobFetchError,
    OutputUploadError,
)

EXPECTED_TOKEN = "EdmPptWk_93e104bb70c17cc560f21e0b46e80bdb"


def _settings(tmp_path, **overrides) -> Settings:
    values = {
        "app_env": "development",
        "backend_base_url": "https://api.example.com/nuSource/api/v1",
        "backend_request_timeout_seconds": 5,
        "get_job_retry_attempts": 1,
        "callback_retry_attempts": 1,
        "temp_root": str(tmp_path / "jobs"),
        "log_root": str(tmp_path / "logs"),
    }
    values.update(overrides)
    return Settings(**values)


def _json_string(request: httpx.Request) -> dict:
    content_type = request.headers.get("content-type", "")
    assert "application/x-www-form-urlencoded" in content_type
    form = parse_qs(request.content.decode(), keep_blank_values=True)
    assert list(form) == ["JSONString"]
    return json.loads(form["JSONString"][0])


MATERIAL = {
    "job_id": 123,
    "asset_id": 456,
    "material_id": 789,
    "material_name": "Week 1 deck",
    "institution_id": 10,
    "institution_name": "Acme",
    "file_url": "https://cdn.example.com/in/a.pptx?Expires=1&Signature=secret",
}

SUCCESS_ENVELOPE = {"code": "success", "data": {"material": MATERIAL}}
EMPTY_ENVELOPE = {"code": "success", "data": {"material": None}}


def test_worker_token_is_the_exact_php_password():
    from api_client.jobs import WORKER_TOKEN

    assert WORKER_TOKEN == EXPECTED_TOKEN


def _client(tmp_path, handler, **overrides) -> BackendClient:
    return BackendClient(
        _settings(tmp_path, **overrides),
        client=httpx.Client(transport=httpx.MockTransport(handler)),
    )


def test_parse_next_maps_material_fields():
    job = parse_next_response(SUCCESS_ENVELOPE)
    assert job.job_id == 123
    assert job.asset_id == 456
    assert job.material_id == 789
    assert job.material_name == "Week 1 deck"
    assert job.institution_id == 10
    assert job.institution_name == "Acme"
    assert "Signature=secret" in job.file_url


def test_parse_next_allows_null_material_id():
    material = {**MATERIAL, "material_id": None}
    job = parse_next_response({"code": "success", "data": {"material": material}})
    assert job.job_id == 123
    assert job.material_id is None


def test_parse_next_empty_queue_when_material_is_null():
    assert parse_next_response(EMPTY_ENVELOPE) is None
    assert parse_next_response({"code": "success", "data": {"material": None}}) is None


def test_parse_next_rejects_non_success_code():
    with pytest.raises(JobFetchError):
        parse_next_response({"code": "error", "data": {"material": MATERIAL}})


def test_parse_next_retry_job_looks_identical_to_a_new_job():
    first = parse_next_response(SUCCESS_ENVELOPE)
    again = parse_next_response(SUCCESS_ENVELOPE)
    assert first == again
    assert first.job_id == 123


def test_get_next_job_sends_only_the_hardcoded_token(tmp_path):
    seen = {}

    def handler(request: httpx.Request) -> httpx.Response:
        seen["method"] = request.method
        seen["path"] = request.url.path
        seen["query"] = dict(request.url.params)
        seen["token_header"] = request.headers.get(TOKEN_HEADER.lower())
        seen["body"] = request.content
        return httpx.Response(200, json=SUCCESS_ENVELOPE)

    job = _client(tmp_path, handler, ppt_worker_token="invented-other-token").get_next_job()
    assert seen["method"] == "GET"
    assert seen["path"].endswith(NEXT_PATH)
    assert seen["query"] == {"token": EXPECTED_TOKEN}
    assert seen["token_header"] == EXPECTED_TOKEN
    assert seen["body"] in (b"", None)
    assert job.job_id == 123


def test_get_next_job_empty_when_material_is_null(tmp_path):
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(200, json=EMPTY_ENVELOPE)

    assert _client(tmp_path, handler).get_next_job() is None


def test_get_next_job_non_200_is_a_fetch_error(tmp_path):
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(500, json={"code": "error"})

    with pytest.raises(JobFetchError):
        _client(tmp_path, handler).get_next_job()


def test_get_next_job_404_is_not_treated_as_empty(tmp_path):
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(404, json={"code": "error"})

    with pytest.raises(JobFetchError):
        _client(tmp_path, handler).get_next_job()


def test_result_success_posts_form_jsonstring_with_token(tmp_path):
    seen = {}

    def handler(request: httpx.Request) -> httpx.Response:
        seen["method"] = request.method
        seen["path"] = request.url.path
        seen["query"] = dict(request.url.params)
        seen["token_header"] = request.headers.get(TOKEN_HEADER.lower())
        seen["payload"] = _json_string(request)
        return httpx.Response(200, json={"code": "success"})

    _client(tmp_path, handler, ppt_worker_token="invented-other-token").send_job_result(
        make_job(job_id=123),
        status=2,
        ispringcloud_link="https://harshit.ispring.com/app/embed-player/abc",
    )
    assert seen["method"] == "POST"
    assert seen["path"].endswith(RESULT_PATH)
    assert seen["query"] == {}
    assert seen["token_header"] == EXPECTED_TOKEN
    assert seen["payload"] == {
        "token": EXPECTED_TOKEN,
        "job_id": 123,
        "status": 2,
        "ispringcloud_link": "https://harshit.ispring.com/app/embed-player/abc",
    }


def test_result_failure_posts_form_jsonstring_with_token(tmp_path):
    seen = {}

    def handler(request: httpx.Request) -> httpx.Response:
        seen["payload"] = _json_string(request)
        return httpx.Response(200, json={"code": "success"})

    _client(tmp_path, handler).send_job_result(
        make_job(job_id=123),
        status=3,
        stage="download",
        error_code="DOWNLOAD_FAILED",
        error_message="The downloaded file is empty.",
    )
    assert seen["payload"] == {
        "token": EXPECTED_TOKEN,
        "job_id": 123,
        "status": 3,
        "stage": "download",
        "error_code": "DOWNLOAD_FAILED",
        "error_message": "The downloaded file is empty.",
    }


def test_result_uses_the_job_id_from_this_claim(tmp_path):
    seen = {}

    def handler(request: httpx.Request) -> httpx.Response:
        seen["payload"] = _json_string(request)
        return httpx.Response(200, json={"code": "success"})

    client = _client(tmp_path, handler)
    client.send_job_result(
        make_job(job_id=123),
        status=2,
        ispringcloud_link="https://example.com/a",
    )
    client.send_job_result(
        make_job(job_id=124),
        status=2,
        ispringcloud_link="https://example.com/b",
    )
    assert seen["payload"]["job_id"] == 124
    assert seen["payload"]["token"] == EXPECTED_TOKEN


def test_backend_client_never_posts_retry(tmp_path):
    urls = []

    def handler(request: httpx.Request) -> httpx.Response:
        urls.append(str(request.url))
        return httpx.Response(200, json=EMPTY_ENVELOPE)

    client = _client(tmp_path, handler)
    client.get_next_job()
    client.send_job_result(
        make_job(job_id=123),
        status=3,
        stage="ispring",
        error_code="ISPRING_FAILED",
        error_message="publish failed",
    )
    assert all(RETRY_PATH not in url for url in urls)
    assert not hasattr(client, "retry_job") or not callable(
        getattr(client, "retry_job", None)
    )


def test_result_non_2xx_raises(tmp_path):
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(500, text="nope")

    with pytest.raises(CallbackError):
        _client(tmp_path, handler).send_job_result(
            make_job(),
            status=3,
            stage="callback",
            error_code="CALLBACK_FAILED",
            error_message="x",
        )


def test_result_stage_mapping():
    assert result_stage("download") == "download"
    assert result_stage("prepare") == "download"
    assert result_stage("powerpoint") == "ispring"
    assert result_stage("powerpoint_open") == "ispring"
    assert result_stage("ispring_publish") == "ispring"
    assert result_stage("upload") == "upload"
    assert result_stage("callback") == "callback"
    assert result_stage("unexpected") == "ispring"


def test_error_code_mapping():
    assert error_code_for(DownloadError("x")) == "DOWNLOAD_FAILED"
    assert error_code_for(ISpringPublishingError("x")) == "ISPRING_FAILED"
    assert error_code_for(OutputUploadError("x")) == "UPLOAD_FAILED"
    assert error_code_for(CallbackError("x")) == "CALLBACK_FAILED"


def test_orphaned_output_is_appended(tmp_path):
    log_root = tmp_path / "logs"
    path = record_orphaned_output(
        log_root,
        make_job(job_id=123, material_id=789),
        UploadResult("ppt/789", "index.html", "https://cdn.example.com/ppt/789/index.html"),
    )
    row = json.loads(path.read_text(encoding="utf-8").splitlines()[0])
    assert row["job_id"] == 123
    assert row["iframe_url"].endswith("index.html")
