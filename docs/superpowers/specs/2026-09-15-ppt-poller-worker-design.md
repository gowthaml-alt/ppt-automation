# PPT Automation Poller Worker — Design

Date: 2026-09-15
Status: Approved for implementation

## 1. Purpose

A long-running Python process on a Windows 365 Cloud PC that converts PowerPoint
files into iSpring HTML5 packages. It is a worker, not a queue owner.

Node.js creates and updates queue records. This worker never connects to MySQL,
never claims or updates queue rows, and never exposes a public job-intake API.

## 2. Architecture

Outbound polling only:

1. `GET` the backend next-job API.
2. If no job, sleep `POLL_INTERVAL_SECONDS` and repeat.
3. If a job is returned, process exactly one PPTX through PowerPoint/iSpring.
4. Upload the HTML5 folder and Playwright-test the public `index.html` URL.
5. `POST` a success or failure callback with the same `queue_id`.
6. Delete local temp files. Then poll again.

The Cloud PC does not need a public IP. FastAPI is not used.

## 3. Constraints

- No MySQL, no queue-table schema, no Python-side claim SQL.
- `queue_id` is an opaque correlation id for GET + callback only.
- One PowerPoint job at a time (the poll loop is single-flight).
- No inbound worker API-key unless the Node API actually requires a header.
- iSpring stays behind an adapter. Default is `not_configured` until probed.
- Playwright tests the published HTML5 in a browser. It does not drive PowerPoint.
- Presigned download URLs are never logged with query strings.

## 4. Internal job model

Mapped from the GET JSON. Extra keys are ignored.

```
queue_id: int            # required (also accepts id / job_id from the API)
material_id: int         # required
ppt_file_url: str        # required, often an S3 presigned GET URL
material_name: str       # optional
institution_name: str    # optional
```

Empty GET: HTTP 204, or 200 with `null` / `{}` / `{"job": null}` / `{"data": null}`.

## 5. Callbacks

Success: `{queue_id, material_id, status: "completed", iframe_url}`

Failure: `{queue_id, material_id, status: "failed", error_message}`

Node.js updates the queue and material. If upload succeeded but the callback
failed, append `LOG_ROOT/orphaned_outputs.jsonl` and keep the uploaded files.

## 6. Pipeline stages

`fetch` → `prepare` → `download` → `powerpoint_open` → `ispring_publish` →
`powerpoint_close` → `output_validate` → `upload` → `browser_test` →
`callback` → `cleanup`

Cleanup always runs in `finally`.

## 7. Known unknowns

Exact Node GET/callback paths and empty-job JSON are configurable
(`BACKEND_GET_JOB_PATH`, `BACKEND_CALLBACK_PATH`). Field aliases are accepted
so a later sample response does not require a redesign.
