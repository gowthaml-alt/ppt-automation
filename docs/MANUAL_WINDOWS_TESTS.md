# Manual Windows integration tests

Do not treat a green `pytest` run as proof that publishing works. These steps run on the Windows 365 Cloud PC with PowerPoint and iSpring installed.

1. Log in as the automation user (not SYSTEM).
2. `python scripts/probe_ispring.py` and keep `logs/ispring-probe.json`.
3. Open PowerPoint manually, load a small PPTX, publish with iSpring by hand, confirm `index.html` plus asset folders.
4. Point `.env` at a non-production Node GET/callback pair or a mock server that returns one job.
5. Set `ISPRING_ADAPTER=not_configured` first: the worker must send a **failure** callback, then clean temp files.
6. After a real adapter is implemented, run one live job:
   - GET returns `queue_id`, `material_id`, presigned `ppt_file_url`
   - Output appears on the CDN at `/ppt/<material_id>/index.html`
   - Playwright loads that URL
   - Node receives `status=completed`
7. Force a download failure (expired presigned URL) and confirm `status=failed` callback.
8. Confirm PowerPoint is not left running after success or failure.
9. Confirm the Cloud PC has no inbound job port requirement.

Until step 6 succeeds, do not claim the automation works.
