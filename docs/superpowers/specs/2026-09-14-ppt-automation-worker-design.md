# PPT Automation Worker — Design (SUPERSEDED)

Date: 2026-09-14
Status: **Superseded** by `docs/superpowers/specs/2026-09-15-ppt-poller-worker-design.md`

This document described a MySQL-claiming FastAPI intake worker. That architecture is incorrect. Do not implement it.

## 1. Purpose

A Python FastAPI service that runs on a dedicated Windows 365 Cloud PC and converts
uploaded PowerPoint files into published iSpring HTML5 packages hosted at an
embeddable iframe URL.

The service is a worker, not an orchestrator. An existing PHP backend owns upload,
job creation, and material state. An existing PHP cron pushes pending jobs to this
service one at a time. There is no Linux worker in this architecture; every PPT is
processed through PowerPoint Desktop and iSpring Suite on Windows.

Expected volume is approximately ten PPTs per day. The design is deliberately sized
for a single machine: no Celery, Redis, RabbitMQ, or Kubernetes.

## 2. Scope

In scope for version one:

- Authenticated job intake, atomic job claiming, and status reporting.
- Streaming download of the input PPT with size, timeout, and integrity limits.
- PowerPoint Desktop automation via pywin32/COM.
- An iSpring publishing adapter behind a replaceable interface.
- Recursive upload of the published folder to storage with correct MIME types.
- HTTP verification of the published entry point.
- Backend notification with the resulting iframe URL.
- Slack notification on failure.
- Temporary file cleanup.

Explicitly out of scope:

- Parallel PowerPoint processing.
- PPT classification logic.
- A retry system.
- Any `current_step`, `worker_type`, `retry_count`, `retry_at`, or `iframe_url`
  column in this worker's job table.
- Browser-based rendering verification (deferred; see section 9.3).

## 3. Architecture

### 3.1 Process model

A single FastAPI application served by uvicorn with exactly one worker process.

This is a correctness constraint, not a performance choice. The "one PPT at a time"
guarantee is enforced by an in-process lock. A second uvicorn worker would create a
second lock and allow two concurrent PowerPoint sessions. The application logs a
warning at startup if it can detect more than one worker, and the README states the
constraint prominently.

### 3.2 Layering

Imports flow in one direction only:

```
api  ->  services  ->  integrations
              \->  utils
```

- `api` handles HTTP concerns: authentication, request validation, status codes.
  It never imports pywin32 and never performs COM calls.
- `services` holds the workflow. `job_service` owns claim and status transitions;
  `job_runner` owns pipeline orchestration.
- `integrations` holds process-boundary code: database, HTTP client, COM.

No module in `api` or `services` imports a Windows-only module at module scope. All
COM access is funnelled through `integrations/windows_com.py`, which guards its own
imports behind a platform check. Consequence: the entire unit test suite imports and
runs on macOS and Linux.

### 3.3 COM thread affinity

COM objects have thread affinity and pywin32 calls block. All PowerPoint and iSpring
work therefore runs on one dedicated worker thread owned by `job_runner`
(`ThreadPoolExecutor(max_workers=1)`), entered through a `com_apartment()` context
manager that pairs `CoInitialize` and `CoUninitialize`. The asyncio event loop
delegates to that thread and is never blocked by a publish that may take 30 minutes.

### 3.4 Execution model

`POST /jobs/process` claims the job synchronously inside the request, schedules the
pipeline as a background asyncio task, and returns `202 Accepted` immediately. The
PHP cron does not hold an HTTP connection open for the duration of a publish.

Background tasks are created with `asyncio.create_task` and their references are held
in a set on application state so they cannot be garbage collected mid-flight. On
application shutdown, an in-flight job is logged as interrupted and PowerPoint is
closed; the job row is marked FAILED with an `interrupted` stage so it is visible
rather than stuck at PROCESSING.

## 4. Concurrency and duplicate protection

Two independent guards, both required.

### 4.1 In-process single slot

A `threading.Lock` wrapped in a `JobSlot` object exposing `try_acquire()` and
`release()`. A `threading.Lock` is used rather than an `asyncio.Lock` because it
offers a genuinely atomic non-blocking acquire; testing `asyncio.Lock.locked()` and
then acquiring is a race, and it is also usable from the COM worker thread.

### 4.2 Database atomic claim

```sql
UPDATE ppt_automation_jobs
SET status = 1,
    updated_at = NOW()
WHERE id = :job_id
  AND status = 0;
```

The affected row count is the authority. Zero rows means the job was not claimable.

### 4.3 Request sequence

Order is significant:

1. Validate `X-Worker-API-Key` using `secrets.compare_digest`. Reject with `401`
   before any other work.
2. Validate the request body via a pydantic schema. Reject with `422`.
3. `slot.try_acquire()`. On failure return `423 Locked` with reason `worker_busy`
   and **do not touch the database**. The job stays PENDING and the next cron tick
   retries it. No hidden queue state exists anywhere.
4. Execute the atomic claim. If `rowcount == 0`, read the row to classify the
   outcome, release the slot, and return `409 Conflict` with one of `not_found`,
   `already_processing`, `already_completed`, `already_failed`.
5. Verify the claimed row's `material_id` equals the request's `material_id`. A
   mismatch means caller and database disagree about the job's identity; mark the
   job FAILED and return `409` rather than risk publishing the wrong material.
6. Schedule the pipeline task and return `202 Accepted`.

Claiming inside the request rather than the background task means a `202` response is
itself proof that the job was claimed exactly once. The background task releases the
slot in a `finally` block.

## 5. Data model

The existing table is used unchanged. This project writes no DDL and no migrations.

```sql
CREATE TABLE ppt_automation_jobs (
    id BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    material_id BIGINT NOT NULL,
    material_name VARCHAR(500) NOT NULL,
    institution_name VARCHAR(500) NOT NULL,
    ppt_file_url TEXT NOT NULL,
    status TINYINT NOT NULL DEFAULT 0,
    error_message TEXT NULL,
    created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
        ON UPDATE CURRENT_TIMESTAMP,
    UNIQUE KEY uk_material_id (material_id),
    INDEX idx_status (status)
);
```

Status values: `0` PENDING, `1` PROCESSING, `2` COMPLETED, `3` FAILED.

`error_message` is a `TEXT` column, so recorded errors are truncated to a safe length
before writing, with the full detail always present in the logs.

### 5.1 State machine

```
PENDING(0) --claim--> PROCESSING(1) --+-- every step succeeded --> COMPLETED(2)
                                      |
                                      +-- any failure ----------> FAILED(3)
```

`COMPLETED` is written only after the backend update returns success. There is no
transition out of `COMPLETED`. Resetting a failed job is a manual operator action
documented in the README, not a code path.

### 5.2 Access layer

MySQL through SQLAlchemy 2.x **Core** (not the ORM) with the PyMySQL driver:
`DATABASE_URL=mysql+pymysql://user:pass@host:3306/dbname`.

A synchronous engine with `pool_pre_ping=True`, invoked from threads. At ten jobs a
day an async driver adds a dependency and buys nothing. The driver name lives in the
URL, so swapping to `mysqlconnector` requires no code change.

All statements use `text()` with bound parameters. No string concatenation of user
input into SQL anywhere.

Repository methods: `get_job`, `claim_pending_job`, `mark_processing`,
`mark_completed`, `mark_failed`, `get_status`.

## 6. Pipeline

Ordered stages, each one a named `stage` value in logs and in `error_message`:

1. `claim` — atomic claim (performed in the request, see section 4.3).
2. `prepare` — create the job directory tree; check free disk against
   `MIN_FREE_DISK_GB`.
3. `download` — stream the input PPT.
4. `powerpoint_open` — start PowerPoint, open the presentation.
5. `ispring_publish` — invoke the publishing adapter.
6. `powerpoint_close` — close the presentation and quit PowerPoint. Runs as soon as
   publishing returns, and again idempotently from the teardown path, so PowerPoint
   is never left holding the file during upload.
7. `output_validate` — validate the published folder locally.
8. `upload` — recursive folder upload to storage.
9. `verify` — HTTP `HEAD` on the uploaded `index.html`.
10. `backend_update` — notify the existing backend with the iframe URL.
11. `complete` — mark the job COMPLETED.
12. `cleanup` — always, in a `finally` block.

### 6.1 Job directories and path safety

```
D:\ppt-automation\jobs\<job_id>\
    input\
    working\
    output\
    logs\
```

The directory name is derived from the integer `job_id` only. `material_name`,
`institution_name`, and the filename component of `ppt_file_url` never influence any
filesystem path. The downloaded file is always named `input\source.pptx` regardless
of the remote filename. Every constructed path is resolved and asserted to be inside
`TEMP_ROOT` before use, which makes path traversal structurally impossible rather
than filtered.

### 6.2 Download requirements

- Streamed with `httpx`, chunked writes, separate connect and read timeouts derived
  from `DOWNLOAD_TIMEOUT_SECONDS`.
- Written to `source.pptx.part` and renamed only after all checks pass.
- Aborted if the running byte total exceeds `MAX_PPT_SIZE_MB`, so a lying or absent
  `Content-Length` cannot fill the disk.
- Verified to exist and be non-empty.
- Verified to begin with the OOXML/ZIP magic bytes `PK\x03\x04`, which catches an
  HTML error page saved with a `.pptx` name.
- URL scheme restricted to `http` or `https`. An optional `DOWNLOAD_ALLOWED_HOSTS`
  allowlist restricts the host; empty means allow-all and the SSRF implication is
  documented in the README.

### 6.3 Output validation

The worker validates the published folder before uploading anything:

- The publish output directory exists and is non-empty.
- `index.html` exists at its root and is non-empty.
- At least one asset subdirectory exists and contains files, guarding against a
  publish that wrote only a stub.
- A file manifest (relative path, size) is recorded in the job log for diagnosis.

The worker uploads the complete directory tree with relative structure preserved. It
never uploads `index.html` alone.

## 7. Replaceable seams

Three boundaries are defined as protocols so the unknown or environment-specific
parts are swappable and independently testable.

### 7.1 iSpring publisher

```python
class ISpringPublisher(Protocol):
    def is_available(self) -> ISpringAvailability: ...
    def publish(self, pptx: Path, output_dir: Path, timeout_s: int) -> Path: ...
```

Selected by `ISPRING_ADAPTER`. Implementations:

- `not_configured` (default) — raises `ISpringNotConfiguredError` with instructions
  pointing at the probe script. This is the shipped default because the real
  publishing mechanism has not yet been verified on the installed version.
- `fake` — used by tests; copies a fixture folder shaped like real iSpring output.
- `vba`, `uia`, `cli` — stubs with a documented contract and no invented API
  surface. Exactly one will be completed once the probe output is available.

No iSpring API, executable name, command-line flag, ProgID, or control identifier is
guessed anywhere in the codebase. `scripts/probe_ispring.py` gathers the facts needed
to complete the integration:

- The PowerPoint `COMAddIns` collection: ProgID, description, connect state, GUID.
- iSpring installation directories and candidate executables with versions.
- Relevant registry entries under the add-in and iSpring keys.
- The ribbon UI Automation tree for the iSpring tab, including control types, names,
  and automation IDs.
- Whether the add-in exposes anything callable through `Application.Run`.

The script writes a report file to `LOG_ROOT` for review. Choosing and completing the
real adapter is a follow-up task gated on that report.

### 7.2 Storage backend

```python
class StorageBackend(Protocol):
    def upload_directory(self, local: Path, prefix: str) -> UploadResult: ...
    def verify(self, entry_url: str) -> bool: ...
```

`UploadResult` carries `output_path` (the storage prefix), `entry_file`
(`index.html`), and `iframe_url` (the absolute HTTP URL).

Selected by `STORAGE_BACKEND`:

- `s3` — boto3 against any S3-compatible endpoint. `S3_ENDPOINT_URL` is configurable
  for non-AWS providers. `S3_PUBLIC_BASE_URL` supplies the CDN origin used to build
  the iframe URL, because a bucket URL and a served URL are not the same thing.
- `http_api` — uploads to an endpoint on the existing backend. The exact contract is
  not yet known; the implementation is a documented stub listing the required facts
  (endpoint path, auth scheme, multipart or archive format, response shape).
- `local_fs` — copies into a directory. Used for development and as the storage test
  double.

MIME types are set explicitly per file extension for `.html`, `.js`, `.css`, `.json`,
`.svg`, `.woff`, `.woff2`, `.ttf`, `.mp4`, `.mp3`, `.wasm`, and images, falling back
to `mimetypes.guess_type`. This is a correctness requirement: iSpring HTML5 output
renders as a blank frame when a CDN serves JavaScript as `application/octet-stream`.

Input download is deliberately not part of this interface. The input URL may be an
arbitrary signed URL unrelated to the output storage provider, so downloading is a
standalone streaming HTTP function.

### 7.3 Backend client

```python
class BackendClient(Protocol):
    def update_material_output(self, job: JobContext, result: UploadResult) -> None: ...
```

Posts to `BACKEND_BASE_URL` + `BACKEND_UPDATE_PATH` with `material_id`, `job_id`,
`output_path`, `entry_file`, and `iframe_url`. Authenticated by a key read from the
environment, sent in a configurable header name. Explicit connect and read timeouts.
No hardcoded URLs or secrets. Non-2xx or a malformed response raises
`BackendUpdateError`.

The existing backend stores `iframe_url`; this worker's table does not.

## 8. Error handling

### 8.1 Hierarchy

All custom exceptions derive from `PptAutomationError`, which carries a `stage`
string and a `user_message`. That single shared attribute is what makes the failure
path uniform — whichever exception propagates, the handler knows what to record in
`error_message`, what to put in the Slack notification, and what to set as the
structured log's `stage` field.

`JobClaimError`, `DownloadError`, `PowerPointAutomationError`,
`ISpringPublishingError`, `ISpringNotConfiguredError`, `ISpringTimeoutError`
(subclass of `ISpringPublishingError`, so a timeout is distinguishable from a fast
failure), `OutputValidationError`, `OutputUploadError`, `BackendUpdateError`,
`DatabaseError`, `CleanupError`, `LowDiskSpaceError`.

### 8.2 Rules

No bare `except:` and no broad `except Exception` without `exc_info=True`. Exactly
two places catch broadly:

- The pipeline's outermost handler, which converts any unexpected exception into a
  FAILED job with stage `unexpected`.
- `slack_service`, which swallows everything by design, because a Slack outage must
  never fail a job that otherwise succeeded.

Every enumerated failure mode has an explicit handler: invalid request,
authentication failure, job not found, job already claimed, download timeout,
download HTTP error, empty or corrupt file, PowerPoint startup failure, PowerPoint
open failure, COM errors, iSpring not installed, publishing failure, publishing
timeout, missing output, upload failure, backend update failure, database failure,
Slack failure, cleanup failure, and low disk space.

COM errors are translated at the boundary: `windows_com.py` catches
`pywintypes.com_error`, extracts the HRESULT and description, and raises
`PowerPointAutomationError`. No `pywintypes` type escapes into the service layer.

### 8.3 Partial success: upload succeeded, backend update failed

The expensive work is already done, so the uploaded output is not rolled back. Before
marking the job FAILED, the worker appends a JSON line to
`LOG_ROOT\orphaned_outputs.jsonl` containing `job_id`, `material_id`, `iframe_url`,
`output_path`, and a timestamp, and embeds the iframe URL in `error_message`. That
file lives outside the job directory, so cleanup cannot destroy it. The README
documents how to replay these records.

### 8.4 Timeouts

Every wait is bounded: `DOWNLOAD_TIMEOUT_SECONDS`,
`POWERPOINT_START_TIMEOUT_SECONDS`, `POWERPOINT_OPEN_TIMEOUT_SECONDS`,
`ISPRING_PUBLISH_TIMEOUT_SECONDS`, `UPLOAD_TIMEOUT_SECONDS`,
`BACKEND_REQUEST_TIMEOUT_SECONDS`, `SLACK_REQUEST_TIMEOUT_SECONDS`. Polling loops
compare against a deadline computed once at entry, never an accumulated sleep count.

## 9. Supporting services

### 9.1 PowerPoint service

Verifies installation without launching PowerPoint (configured path plus COM
registration). Starts PowerPoint, opens a presentation with automation-friendly
arguments to suppress prompts where the object model allows, closes the presentation,
and quits the application. Close and quit are idempotent and safe to call from a
`finally` block after a failed open.

Orphaned `POWERPNT.exe` detection uses `psutil` and only terminates a process when
the job slot is free and the process predates a configured age threshold, gated by
`POWERPOINT_KILL_ORPHANS`. The service assumes a dedicated interactive Windows
automation user, not the SYSTEM account; the README explains why and what breaks
otherwise.

### 9.2 Slack service

Posts a webhook message for failed jobs containing job ID, material ID, material
name, institution name, failure stage, and error message. Never raises. The webhook
URL is never logged, not even partially.

### 9.3 Verification depth

Verification is a local manifest check followed by an HTTP `HEAD` (falling back to a
ranged `GET`) on the uploaded `index.html`. Browser-based rendering verification with
Playwright is deliberately deferred: it would add a browser dependency to the worker
for a check better performed once during manual Windows integration testing, which is
where it appears in the manual test document.

### 9.4 Cleanup service

Runs in a `finally` block inside its own `try/except`. On success it deletes the
entire job directory after the upload and backend update have both succeeded. On
failure it copies the per-job log to `LOG_ROOT` first, then deletes the job directory
unless `KEEP_FAILED_JOB_FILES` is true (default `false`).

This resolves an ambiguity in the requirements, which asked both that temporary files
be deleted on failure and that nothing be deleted before a successful upload. The
default favours disk hygiene; the flag exists for debugging a failing publish.

A `CleanupError` is logged as a distinct event and never replaces the original
exception. The original error is captured before cleanup runs and recorded afterwards.

## 10. Logging

Standard library `logging` with a JSON formatter and a `contextvars`-based context
that binds `job_id`, `material_id`, and `stage` onto every record emitted during a
job. No structlog or other logging framework.

Every stage logs start and completion with a `duration_ms` field. Failures log the
exception with `exc_info=True`.

A redaction filter strips query strings from any logged URL and masks configured
secret values. Signed URLs therefore appear as scheme, host, and path only.
Passwords, API keys, and access tokens are never logged.

Handlers: a rotating file handler under `LOG_ROOT`, a console handler, and a per-job
file handler writing into the job's `logs\` directory.

## 11. Configuration

`pydantic-settings`, loaded from the environment and an optional `.env` which is never
committed.

Beyond the variables named in the requirements, the following are added:
`STORAGE_BACKEND`, `ISPRING_ADAPTER`, `S3_BUCKET`, `S3_REGION`, `S3_ENDPOINT_URL`,
`S3_ACCESS_KEY_ID`, `S3_SECRET_ACCESS_KEY`, `S3_KEY_PREFIX`, `S3_PUBLIC_BASE_URL`,
`BACKEND_UPDATE_PATH`, `BACKEND_API_KEY_HEADER`, `DOWNLOAD_ALLOWED_HOSTS`,
`KEEP_FAILED_JOB_FILES`, `MIN_FREE_DISK_GB`, `POWERPOINT_KILL_ORPHANS`,
`POWERPOINT_ORPHAN_MAX_AGE_SECONDS`, `UVICORN_WORKERS`, `LOG_LEVEL`.

Startup validation refuses to start when `APP_ENV=production` and any of the following
hold, because each would let the service run in a state that looks healthy while
being wrong:

- `WORKER_API_KEY` is empty or still the placeholder value.
- `STORAGE_BACKEND=local_fs`, which would write published output to a local
  directory that nothing serves, producing an iframe URL that cannot resolve.

## 12. API

### `GET /health`

Always returns HTTP 200 with a `status` of `ok` or `degraded` and a per-check
breakdown, so a monitor can read the reason rather than only observing a failure.
Checks: FastAPI liveness, PowerPoint availability, iSpring availability, temp
directory writability (a real write-and-delete probe), free disk space against
`MIN_FREE_DISK_GB`, and database connectivity. Reveals no secrets or credentials.

### `POST /jobs/process`

Requires `X-Worker-API-Key`. Body: `job_id`, `material_id`, `material_name`,
`institution_name`, `ppt_file_url`, all required. Responses: `202` accepted, `401`
bad key, `409` not claimable, `422` invalid body, `423` worker busy.

The `409` body carries a `reason` of `not_found`, `already_processing`,
`already_completed`, `already_failed`, or `material_mismatch` (section 4.3, step 5).

### `GET /jobs/{job_id}`

Returns the current status from the database. Requires the same API key.

## 13. Testing

The full unit suite runs on macOS or Linux with no PowerPoint, no iSpring, and no
MySQL installed.

`pytest` with `httpx` `ASGITransport` for API tests. Test doubles for all six
boundaries: `FakePowerPointService` (records open/close/quit calls without COM),
`FakeISpringPublisher` (copies a fixture folder shaped like real iSpring output),
`LocalFsStorageBackend` (doubles as the storage fake), and fakes for the backend
client, Slack, and the job repository. Every boundary is injected into `job_runner`
through its constructor, which is what makes the pipeline testable off Windows.

| Area | Approach |
| --- | --- |
| Health endpoint | Fakes for each probe, asserting `ok` and `degraded` shapes |
| Request validation | Missing and wrong-typed fields produce `422` |
| API-key auth | Missing, wrong, and correct keys |
| Job claiming | SQL text and bound parameters asserted against a mocked connection |
| Duplicate submission | Two concurrent ASGI requests: exactly one `202`, one `423`, pipeline invoked once |
| Status transitions | Outcome mapping from `rowcount` and current status |
| Download failure | Timeout, non-2xx, empty body, oversize, bad magic bytes |
| Cleanup execution | Runs on success and failure; honours `KEEP_FAILED_JOB_FILES` |
| Error handling | Each stage raises its own exception type and records the right stage |
| Storage failure | Upload raises `OutputUploadError`; job FAILED, no backend call |
| Backend update failure | Job FAILED and an `orphaned_outputs.jsonl` line is written |

Optional real-MySQL tests are marked `@pytest.mark.mysql` and skipped unless a
database URL is supplied.

`docs/MANUAL_WINDOWS_TESTS.md` covers the Windows-only path: running the probe
script, a COM smoke test, the first real publish, DPI and interactive-session
requirements, and recovery drills.

## 14. Repository layout

The structure from the requirements, plus:

- `app/exceptions.py` — the exception hierarchy.
- `app/services/job_runner.py` — pipeline orchestration, keeping `job_service.py`
  focused on claim and status.
- `app/integrations/storage_backends/` — the three storage implementations.
- `docs/` — the manual Windows test guide and this design document.

## 15. Dependencies

`fastapi`, `uvicorn[standard]`, `pydantic`, `pydantic-settings`, `httpx`,
`sqlalchemy`, `pymysql`, `boto3`, `psutil`, `python-dotenv`; Windows-only extras
`pywin32` and `pywinauto` declared with an environment marker so installs succeed on
macOS and Linux for development and testing. Test extras: `pytest`, `pytest-asyncio`,
`pytest-cov`.

## 16. Known unknowns

These are genuine external unknowns, not deferred decisions. Each has a concrete
resolution path.

1. **The iSpring publishing interface.** Resolved by running
   `scripts/probe_ispring.py` on the Cloud PC and reviewing its report. Until then
   `ISPRING_ADAPTER=not_configured` is the default and the pipeline fails loudly at
   the publish stage rather than silently producing nothing.
2. **The existing backend's update contract.** Endpoint path, authentication scheme,
   request body shape, and response shape are configurable, and the required facts
   are listed in the README so the values can be filled in without code changes.
3. **The storage/CDN topology.** The `s3` backend covers S3-compatible storage with a
   configurable public base URL; the `http_api` backend is a documented stub for the
   case where the existing backend owns uploads.

Nothing in this design claims the end-to-end automation works. That claim requires a
successful real publish on the Windows machine, which is the exit criterion of
`docs/MANUAL_WINDOWS_TESTS.md`.
