# PPT Automation Worker

Windows 365 worker that publishes PowerPoint files through desktop PowerPoint and iSpring Suite, then reports the public HTML5 URL back to the Node.js backend.

This process **does not connect to MySQL**. It **does not claim or update queue rows**. Node.js owns the queue. The worker only:

1. `GET` the next job from the backend
2. Download the PPTX (the URL may be an S3 presigned GET URL)
3. Publish via PowerPoint/iSpring
4. Upload the HTML5 folder
5. Playwright-test the public `index.html`
6. `POST` a success or failure callback using the same `queue_id`

The Cloud PC does **not** need a public IP. There is no inbound job API.

## Architecture

```
Node.js queue  --GET-->  Windows poller  --POST callback-->  Node.js
                              |
                              +--> PowerPoint + iSpring
                              +--> S3/CDN upload
                              +--> Playwright (HTML5 only)
```

One OS process. One PowerPoint job at a time. The poll loop does not fetch the next job until the current job has finished (including cleanup).

## Windows requirements

- Windows 365 Cloud PC
- Microsoft PowerPoint Desktop
- iSpring Suite
- Python 3.11+
- An interactive automation user (not `SYSTEM` — COM/PowerPoint/iSpring fail under SYSTEM)
- Playwright browsers (`playwright install chromium`) after pip install

## Installation

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
pip install -r requirements-dev.txt
playwright install chromium
copy .env.example .env
```

Fill in `BACKEND_BASE_URL`, storage credentials, and paths. Leave `ISPRING_ADAPTER=not_configured` until `scripts\probe_ispring.py` has been run on this machine.

## Environment variables

See `.env.example`. Important:

| Variable | Purpose |
| --- | --- |
| `BACKEND_BASE_URL` | Node.js origin |
| `BACKEND_GET_JOB_PATH` | GET next job |
| `BACKEND_CALLBACK_PATH` | POST result |
| `BACKEND_API_KEY` | Optional outbound header; unset means no auth header |
| `POLL_INTERVAL_SECONDS` | Sleep when GET returns no job |
| `TEMP_ROOT` / `LOG_ROOT` | Local job files and worker logs |
| `STORAGE_BACKEND` | `s3` in production |
| `S3_PUBLIC_BASE_URL` | Origin used to build `iframe_url` |
| `ISPRING_ADAPTER` | `not_configured` until probed |

There is no `DATABASE_URL` and no `WORKER_API_KEY`.

## How to start

```bat
python main.py
```

or `python run.py`. The process polls forever until SIGINT/SIGTERM.

## Local laptop test (no Node cron)

Give a PPTX on the Windows machine. This does **not** GET a job from Node. It copies the file, optionally opens PowerPoint so you can see the window, publishes (fake iSpring by default), uploads to `LOCAL_STORAGE_ROOT`, then serves the folder on localhost and checks that `index.html` loads.

```bat
copy .env.example .env
```

In `.env` for the laptop:

```bat
APP_ENV=development
BACKEND_BASE_URL=http://127.0.0.1:9
TEMP_ROOT=C:\ppt-automation\jobs
LOG_ROOT=C:\ppt-automation\logs
STORAGE_BACKEND=local_fs
LOCAL_STORAGE_ROOT=C:\ppt-automation\cdn
LOCAL_STORAGE_PUBLIC_BASE_URL=http://127.0.0.1/cdn
MIN_FREE_DISK_GB=1
ISPRING_ADAPTER=fake
```

Then:

```bat
python scripts\process_local.py --pptx "C:\test\sample.pptx" --queue-id 101 --material-id 5001 --material-name "Biology" --institution-name "ABC College"
```

Watch the console. On Windows, PowerPoint should open unless you pass `--skip-powerpoint`.

Success prints `LOCAL RUN PASSED` plus `local_index=` and `checked_url=`. Open `C:\ppt-automation\cdn\ppt\5001\index.html` in a browser.

Add `--playwright` to also load the page in Chromium. Add `--real-ispring` only after the iSpring adapter is implemented.

**This default path uses fake iSpring HTML.** It proves copy → upload → page check. It does not convert your slides until `--real-ispring` works.

Logs: `C:\ppt-automation\logs\worker.log`. Copy failures from that file back to the Mac.

## Local iSpring Cloud test (macOS Chrome)

Playwright can drive the **iSpring Cloud website** in Google Chrome. It cannot drive the PowerPoint / iSpring Suite desktop ribbon on macOS. This script does not start the poller and does not change any cloud content.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python scripts/test_ispring_cloud.py
```

Chrome opens visibly with a persistent profile under `test-artifacts/chrome-profile`. If a login, SSO, or MFA page appears, complete it in Chrome, then press Enter in the terminal. The script searches for `new ppt migration`, opens the first `demo academy` material, prints the title and URL, and writes screenshots plus a Playwright trace under `test-artifacts/ispring-cloud/<timestamp>/`.

## GET job mapping

The client accepts a JSON object (or `{ "job": { ... } }`) with:

- `queue_id` (aliases: `id`, `job_id`)
- `material_id`
- `ppt_file_url` (aliases: `pptx_url`, `file_url`, `ppt_url`)
- optional `material_name`, `institution_name`

Empty queue: HTTP 204, 404, empty body, `{}`, or `{ "job": null }`.

## Callbacks

Success:

```json
{
  "queue_id": 101,
  "material_id": 5001,
  "status": "completed",
  "iframe_url": "https://cdn.example.com/ppt/5001/index.html"
}
```

Failure:

```json
{
  "queue_id": 101,
  "material_id": 5001,
  "status": "failed",
  "error_message": "iSpring publishing failed"
}
```

Node.js updates the queue record. If upload succeeded but the callback failed, the worker appends `LOG_ROOT/orphaned_outputs.jsonl` and keeps the uploaded files.

## PowerPoint / iSpring

PowerPoint is automated with pywin32/COM (`powerpoint/`). iSpring lives only in `publisher/ispring.py`. The shipped adapters are `not_configured` (default, fails loudly), `fake` (unit tests), and unverified stubs (`vba`, `uia`, `cli`). Run `python scripts/probe_ispring.py` on the Cloud PC before implementing a real adapter. Do not assume a CLI.

Playwright opens the **uploaded** `iframe_url`. It does not click the PowerPoint UI.

## Windows startup

Preferred: Task Scheduler at logon of the automation user, running `python main.py`, or NSSM as that user. Keep an interactive session. Auto-logon is the usual Cloud PC pattern for desktop COM.

## Troubleshooting PowerPoint COM

- Do not run as SYSTEM
- Confirm PowerPoint opens manually in the same user session
- RPC_E_SERVERFAULT / RPC_E_CALL_REJECTED often mean a dialog is blocking; the worker cannot click unknown UI
- Orphaned `POWERPNT.EXE` after a crash: close it before the next job

## Troubleshooting iSpring

- Run the probe script and keep `ispring-probe.json`
- Until an adapter is implemented, jobs fail at `ispring_publish` and send a failure callback
- A successful manual publish from the same account is the baseline

## Cleanup

Temp files under `TEMP_ROOT\<queue_id>\` are deleted in `finally` after success or failure. Set `KEEP_FAILED_JOB_FILES=true` only while debugging. Job logs are copied to `LOG_ROOT\jobs\<queue_id>\` on failure before deletion.

## Resetting a failed job

This worker cannot reset queue rows. Re-queue or reset status in Node.js. The next GET will return the job again.

## Security

- Never log presigned URL query strings
- Optional backend API key is outbound-only
- Download host allowlist is empty by default so S3 presigned hosts work; set `DOWNLOAD_ALLOWED_HOSTS` if you can pin them
- The worker does not listen for public job submissions

## Production checklist

- [ ] `APP_ENV=production`
- [ ] `BACKEND_BASE_URL` and paths match the live Node APIs
- [ ] `STORAGE_BACKEND=s3` with `S3_PUBLIC_BASE_URL`
- [ ] iSpring adapter verified on the installed version
- [ ] Playwright chromium installed
- [ ] Task Scheduler / NSSM running as the interactive user
- [ ] Slack webhook optional
- [ ] Manual publish on the Cloud PC succeeded once (`docs/MANUAL_WINDOWS_TESTS.md`)

## Tests

```bat
pytest -v
```

Unit tests mock Node HTTP, PowerPoint, iSpring, storage, Slack, and Playwright. They run on macOS/Linux. They do **not** prove a real iSpring publish.
