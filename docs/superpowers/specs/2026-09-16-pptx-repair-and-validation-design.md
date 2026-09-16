# PPTX Package Validation and Silent Repair

Date: 2026-09-16
Status: Approved for implementation

## 1. Purpose

Unattended Windows publishing must not stop on PowerPoint’s Repair dialog.
The worker validates the downloaded package first, opens a **working copy**
with `Presentations.Open2007(..., OpenAndRepair=True)`, saves that working
copy if repair dirtied it, and continues publishing. The original download
is never overwritten. Password-protected files and Protected View are
hard failures. This pass does **not** click Repair (or any other dialog)
with UI Automation.

## 2. Stages

`download` (package checks) → copy to working → `powerpoint_open`
(Open2007 + timeout + save) → existing publish/upload path.

## 3. Preserve the original

| Path | Role |
| --- | --- |
| `input/source.pptx` | Original download or local copy. Read-only after validation. Never saved over. |
| `working/source.pptx` | Working copy. PowerPoint opens and, if needed, saves repair here. iSpring publishes this path. |

Copy happens at the start of `powerpoint_open`, after download validation
succeeds. `shutil.copy2` from `input/source.pptx` to `working/source.pptx`.

## 4. Download validation

After the existing ZIP magic-byte check (`PK\x03\x04`):

1. Open as a zip. If that fails → invalid package.
2. If the zip contains `EncryptionInfo` or `EncryptedPackage` (any folder
   prefix) → password-protected. Do not open PowerPoint.
3. Require `[Content_Types].xml` and `ppt/presentation.xml`. If either is
   missing → missing parts.

These checks run for HTTP download and for `copy_local_pptx`.

## 5. PowerPoint open (Windows)

`POWERPOINT_OPEN_TIMEOUT_SECONDS` (default 120) wraps the **entire** open
body: `Open2007`, Protected View detection, and Save. A watchdog thread
starts before that body and, on expiry, terminates PowerPoint then raises
the timeout error.

Open call:

```
Presentations.Open2007(
    FileName=working_copy,
    ReadOnly=False,
    Untitled=False,
    WithWindow=True,
    OpenAndRepair=True,
)
```

`DisplayAlerts = ppAlertsNone` (value `1`). Window stays visible so iSpring
can load. Do not call `Presentations.Open`. Do not click Repair.

After a successful open:

- If `ProtectedViewWindows.Count > 0`, close those windows **without**
  enabling editing and fail Protected View.
- If COM reports a password error, fail password-protected.
- If `presentation.Saved` is false, `Save()` to the working copy.
- Leave `input/source.pptx` unchanged.

## 6. Cleanup on open timeout or failure

Always, in this order, best-effort:

1. Close the presentation if one was obtained.
2. `Application.Quit`.
3. Terminate `POWERPNT.EXE` (`taskkill /F /IM POWERPNT.EXE` on Windows).

Pipeline teardown (`close`/`quit`) still runs in `finally`.

## 7. Logging (`stage=powerpoint_open`)

Every attempt logs:

| Field | Meaning |
| --- | --- |
| `repair_attempted` | Always `true` when `Open2007` is invoked with `OpenAndRepair=True`. |
| `saved` | `true` only if this open called `Save()` successfully. |
| `open_result` | `success`, `timeout`, `password`, `protected_view`, `damaged`, or `save_failed`. |

## 8. Exact `user_message` values

Callers (Node callback, Slack) must receive these strings unchanged.

| Condition | Exception | `stage` | `user_message` |
| --- | --- | --- | --- |
| Not ZIP / HTML named `.pptx` | `DownloadError` | `download` | The downloaded file is not a valid PowerPoint file. |
| ZIP magic present but not a readable package | `DownloadError` | `download` | The PowerPoint file is not a valid package. |
| Readable zip missing `[Content_Types].xml` or `ppt/presentation.xml` | `DownloadError` | `download` | The PowerPoint file is missing required package parts. |
| OOXML encryption parts, or COM password error | `DownloadError` at download; `PowerPointAutomationError` at open | `download` or `powerpoint` | The PowerPoint file is password-protected. |
| Protected View window present | `PowerPointAutomationError` | `powerpoint` | PowerPoint opened the file in Protected View. |
| `Open2007` missing, COM failure, or unrepairable file | `PowerPointAutomationError` | `powerpoint` | The PowerPoint file is damaged and could not be opened. |
| Watchdog fires for `POWERPOINT_OPEN_TIMEOUT_SECONDS` | `PowerPointAutomationError` | `powerpoint` | Opening the PowerPoint file timed out. |
| Repair left the presentation dirty and `Save()` failed | `PowerPointAutomationError` | `powerpoint` | The repaired PowerPoint file could not be saved. |

Password at download wins when encryption parts are visible in the zip.
Password at open covers COM/password dialogs that the zip check cannot see.
Do not type a password. Do not click Protected View “Enable Editing”.

## 9. Out of scope

- UI Automation to click Repair, password, or Protected View.
- Writing the repaired bytes back to `input/source.pptx`.
- Changing the iSpring adapter.

## 10. Tests (must exist)

Valid PPTX package; invalid ZIP; missing PPTX parts; password-protected
package; pipeline preserves original and publishes the working copy;
`Open2007` is called with `OpenAndRepair=True`; repairable file is saved
to the working copy; unrepairable COM error; Protected View; open timeout
terminates PowerPoint; Save failure.

Real Repair-dialog suppression, Protected View from a Mark-of-the-Web
file, and `taskkill` of a live `POWERPNT.EXE` are Windows-only and are
not claimed from macOS unit tests.
