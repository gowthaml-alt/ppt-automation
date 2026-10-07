# PPT Automation Worker

Windows 365 worker that publishes PowerPoint files to iSpring Cloud, then
reports the share URL back to the PHP queue APIs.

This process **does not connect to MySQL**. It **does not call retry**. PHP owns
the queue. The worker only:

1. `GET /pptupdate/ispringcloud/next?token=...`
2. Download `file_url` promptly (it is a short-lived signed URL)
3. Publish to iSpring Cloud under `job_id` (cover title is `material_name`)
4. `POST /pptupdate/ispringcloud/result` as form `JSONString` with status `2` or `3`

The Cloud PC does **not** need a public IP. There is no inbound job API.

## PHP APIs

Base: `{API_HOST}/nuSource/api/v1`

This API is form-style. GET uses a query string. POST fields go in `JSONString`
(a form field whose value is a JSON object).

Token: `EdmPptWk_93e104bb70c17cc560f21e0b46e80bdb` (hardcoded, not from env).
Sent as `token` on every call. Optional extra header: `X-PPT-WORKER-TOKEN`.

| Call | Route | Who |
| --- | --- | --- |
| Pull work | `GET /pptupdate/ispringcloud/next?token=...` | this worker |
| Report result | `POST /pptupdate/ispringcloud/result` (`JSONString`) | this worker |
| Requeue a failure | `POST /pptupdate/ispringcloud/retry` | ops only — **never called here** |

Queue statuses: `0` queued, `1` in progress (set when PHP handed you the job),
`2` success, `3` failed.

`job_id` is `ppt_ispring_queue.id`. It is the iSpring presentation name. Material
names are not unique; `job_id` is unique per queue row. A later retry of the
same row reuses that `job_id` and overwrites the existing presentation.

## Architecture

```
PHP queue  --GET /next-->  Windows poller  --POST /result-->  PHP
                                |
                                +--> PowerPoint + iSpring Suite
                                +--> iSpring Cloud share URL
```

One OS process. One job at a time. Never `GET /next` again until the current
job has `POST /result`.

### Where a deck goes

`ISPRING_PARENT_FOLDERS` is now one folder: **PPT migration New**. The order
is:

1. Ask the library for the institution's folder under **PPT migration New**.
   Found - publish there.
2. Not found - read `config/institution_folders.json`. It says which parent
   that institution's folder is under (most of the older ones are under
   **PPT Migration**). Ask the library there and publish into it.
3. Not in either - create the folder under **PPT migration New**, publish,
   and write the new entry into the JSON.

So only one branch of the picker is ever opened, and the map keeps itself up
to date. The file was seeded from the October library listing: 463 entries,
338 under PPT Migration and 125 under PPT migration New. Folders sitting in
the account's default project were left out on purpose - publishing into the
default project is what put decks in "Edmingle Learning Content".

Set `ISPRING_FOLDER_MAP` to point the map somewhere else.

### Private servers

Six institutions run their own copy of the backend, so their queue is on their
own host, not the shared one. Same two routes, same token. The worker polls
them itself:

- shared queue: every `POLL_INTERVAL_SECONDS` (30s)
- private queues: one round every `PRIVATE_POLL_INTERVAL_SECONDS` (600s)

A round walks the servers in order and empties each one before moving to the
next, up to `PRIVATE_MAX_PER_ROUND` decks per server. Whatever is left waits
for the next round. A server that is down is logged and skipped; it never
stops the round or the shared queue.

The result always goes back to the server the job came from — `Job.base_url`
carries it. Post a private server's result to the shared backend and that row
stays in progress for ever.

Hosts are in `Settings.private_base_urls`, copied from
`SupportPanelCatalog.php`. Each private server needs the ppt routes deployed
and its own Slack webhook in `Constants.php`, or its decks publish with no
notification.

## Windows requirements

- Windows 365 Cloud PC
- Microsoft PowerPoint Desktop
- iSpring Suite
- Python 3.11+
- An interactive automation user (not `SYSTEM`)
- Chrome started with `scripts\start_ispring_chrome.cmd` and signed in once

## Installation

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
pip install -r requirements-dev.txt
copy .env.example .env
```

Fill in `BACKEND_BASE_URL` (`{API_HOST}/nuSource/api/v1`). The worker token is a
string constant — do not set it in `.env`.

## How to start

```bat
python main.py
```

or `python run.py`. The process polls forever until SIGINT/SIGTERM.

Idle: if HTTP is not 200, or JSON `code` is not `success`, log, wait, and
`GET /next` again. If `data.material` is null, the queue is empty — wait 10–30s
and poll again.

## Result payloads

`POST /pptupdate/ispringcloud/result` with
`Content-Type: application/x-www-form-urlencoded`.

Success (`status` must be `2`):

```
JSONString={"token":"EdmPptWk_93e104bb70c17cc560f21e0b46e80bdb","job_id":123,"status":2,"ispringcloud_link":"<iframe or share URL>"}
```

Failure (`status` must be `3`):

```
JSONString={"token":"EdmPptWk_93e104bb70c17cc560f21e0b46e80bdb","job_id":123,"status":3,"stage":"download|ispring|upload|callback","error_code":"SHORT_CODE","error_message":"reason"}
```

After `POST /result`, go back to `GET /next`. Do not retry the same job here.

## Local iSpring Cloud test (manual)

```bat
python scripts\publish_material.py --url "https://...signed..." --material-name "Week 1 deck" --material-id 789 --job-id 123 --institution-name "Acme"
```

This does **not** call the PHP queue. Use it to prove PowerPoint + iSpring Cloud
on the machine.

## Cleanup

Temp files under `TEMP_ROOT\cloud-<job_id>\` are deleted after success or
failure. Set `KEEP_FAILED_JOB_FILES=true` only while debugging.

## Resetting a failed job

This worker cannot reset queue rows. Ops calls `POST /pptupdate/ispringcloud/retry`
with `JSONString={"token":"...","job_id":123}`. The next `GET /next` will return
that same `job_id` with a fresh `file_url`. Treat it as a normal job and overwrite
the existing iSpring presentation.

## Security

- Never log signed URL query strings
- Token is outbound-only as `token` (query or JSONString) plus optional `X-PPT-WORKER-TOKEN`
- Download host allowlist is empty by default so signed hosts work
- The worker does not listen for public job submissions

## Tests

```bat
pytest -v
```

Unit tests mock PHP HTTP, PowerPoint, iSpring, Slack, and Playwright. They run
on macOS/Linux. They do **not** prove a real iSpring publish.
