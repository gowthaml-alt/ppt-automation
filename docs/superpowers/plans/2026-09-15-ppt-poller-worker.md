# PPT Poller Worker Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Replace the MySQL/FastAPI intake worker with an outbound poller that GETs jobs from Node.js, publishes via PowerPoint/iSpring adapters, uploads HTML5, Playwright-tests the iframe URL, and POSTs a callback.

**Architecture:** Single long-running Python process. No FastAPI, no SQLAlchemy, no queue table. HTTP client talks to Node. Pipeline collaborators are constructor-injected so unit tests never need Windows, PowerPoint, iSpring, or MySQL.

**Tech Stack:** Python 3.11+, pydantic-settings, httpx, boto3, psutil, playwright (browser check only), pywin32/pywinauto on Windows.

## Global Constraints

- Spec: `docs/superpowers/specs/2026-09-15-ppt-poller-worker-design.md`
- No MySQL. No `DATABASE_URL`. No claim SQL. No inbound `POST /jobs/process`.
- Do not add API-key auth unless `BACKEND_API_KEY` is set (optional outbound header).
- Paths come from integer `queue_id` only. Downloaded file is always `input/source.pptx`.
- Never log URL query strings, API keys, or webhook URLs.
- Default `ISPRING_ADAPTER=not_configured`. Do not invent iSpring APIs.
- Full unit suite must pass on macOS/Linux without PowerPoint, iSpring, or Playwright browsers.
- Do not claim end-to-end publish works until Windows manual tests pass.

## File structure

| Path | Responsibility |
| --- | --- |
| `config/settings.py` | Environment settings |
| `utils/exceptions.py` | Staged errors |
| `utils/logging_config.py` | JSON logs, `queue_id` context, URL scrub |
| `utils/paths.py` | Job directories |
| `utils/validators.py` | PPT URL + magic-byte checks |
| `utils/http.py` | Shared httpx client |
| `utils/cleanup.py` | Temp deletion |
| `api_client/jobs.py` | `get_next_job`, `send_job_result` |
| `downloader/pptx.py` | Stream download |
| `validator/html5.py` | Local output checks |
| `storage/backend.py` | Upload factory + MIME map |
| `powerpoint/service.py` | COM wrapper (no iSpring) |
| `publisher/ispring.py` | Publisher adapters |
| `browser_test/playwright_check.py` | Public iframe check |
| `worker/pipeline.py` | Stage orchestration |
| `worker/loop.py` | Poll loop, single-flight |
| `main.py` | Process entry |
| `scripts/probe_ispring.py` | Windows fact-gatherer |

Remove the old `app/` FastAPI+MySQL package after the new suite passes.

---

### Task 1: Core + HTTP job client + pipeline (this session)

Implement the new packages, tests, README, and delete obsolete intake code. Wire fakes so `pytest` proves: empty GET sleeps; a job is processed once; success/failure callbacks fire; cleanup runs; MySQL and inbound FastAPI are gone.

- [x] Spec written
- [x] Implement packages and tests
- [x] Delete `app/` intake
- [x] `pytest` passes (71 tests)
