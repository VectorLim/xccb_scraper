# xCCB Flawless Submission reporting

Python pipeline for loading the Forum Reviewers workbook, filtering submission rows, reading Actions through Windows Integrated Authentication, and producing an auditable Excel report. All network operations are GET requests. No xCCB records, report subscriptions, or deployments are modified.

## Setup and execution

Python 3.10+ on Windows, with access to the corporate network and xCCB under the account running the job:

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
.\.venv\Scripts\python -m pip install pytest
Copy-Item config.example.json config.json
```

Beautiful Soup handles HTML parsing and serialization with Python's built-in `html.parser`; openpyxl handles Excel. No browser runtime is required.

The selection rules are: Reviewing Forum contains `admin`; Is FYI? equals `No`; Phase does not contain `WLA` (case-insensitive); Document Submitted is in the reporting month (the current Malaysia month by default); and Reviewer(s) is blank or contains at least one exact name from the allowed reviewer list. Multiple reviewer identities in a source cell are separated by semicolons or newlines; commas are kept within names such as `Surname, Given name`. An empty allowed list accepts only blank Reviewer(s) cells. Every non-WLA phase is accepted, including blank phase cells. Text comparisons ignore case and surrounding/repeated whitespace. Author Manager is not an eligibility filter: qualifying rows are sorted into adjacent manager groups, and a `By Author Manager` sheet summarizes each group's Yes/No/unresolved counts. Missing managers are retained under `(Not specified)` in the summary. Source manager cells are preserved.

One **listed names** list drives both checks from the requestor's guide: Reviewer(s) must contain a listed name, and any Action whose `Assigned By` is a listed name makes the row `No`. The list lives in `reviewer_names.txt` beside the executable (working directory for source/CLI runs), one name per line; when that TXT exists it overrides both `reviewer_names` and `assigned_by_names` in config.json. Without the TXT, the two config lists are used as before. Exact names match after case and whitespace normalization; commas are part of names. Blank lines in the TXT are ignored and duplicate names are collapsed without changing the first spelling. UTF-8, including a Notepad BOM, is supported. `action_scope` remains `unconfirmed` until whole-document Actions history is approved. Submission/phase-specific scopes are not implemented. Legacy `reviewer_separator`, `manager_organizations`, `allowed_organizations`, and `phases` settings are ignored with a warning. The phase rule is fixed to accept everything except values containing `WLA`.

Run from the folder containing `config.json`. Supply the source workbook and a new output filename. Add `--live` to capture fresh Actions for each unique qualifying document in `outputs/evidence`; omit it to use previously captured evidence:

```powershell
.\.venv\Scripts\python -m xccb_report 'ForumReviewers (Includes Metadata).xlsx' outputs\Flawless_Submission.xlsx --live
```

There are no subcommands or configuration/output-path flags. `--live` is the only operational option; `--help` shows usage.

## Windows dashboard and executable

Double-click `dist/xCCB_Report.exe`. It includes Python, Tkinter, openpyxl, Beautiful Soup and timezone data; no separate Python installation is required on the target machine. Live access still requires Windows, corporate network access and an authorized Windows account.

1. Enter the **Listed names** in the always-visible, scrollable column on the left (one name per line). The same list is used for the Reviewer(s) filter and the Actions `Assigned by` check.
2. Click **Workbook…** to select the Forum Reviewers `.xlsx` workbook. Select its reporting **Month** as `YYYY-MM` beside Refresh; it defaults to the current Malaysia month. Older workbooks require the corresponding earlier month.
3. Click **Refresh**. Summary cards and the **By Author Manager** tab update as each document finishes. When the run completes, the **Submissions** tab lists every qualifying row with its Flawless result, notes and original `Document WS URL`. Filter by result or search; double-click a row (or **Open link**) to open the xCCB workspace, or **Copy link**. Double-clicking a manager shows that manager's submissions.
4. Click **Export Excel…** to save the finalized report workbook for publishing or emailing.
5. Open **Settings** to change the evidence location, live capture or whole-document scope confirmation. Confirm whole-document scope only when appropriate; this does not attest completeness of individual captures.

The dashboard generates the report in a temporary workbook and removes it after reading; a copy is written only when you use **Export Excel…**.

In **Submissions**, click any column heading to open its Excel-style value filter. Search the list, check the values to keep, then click **Apply**. Filters on multiple columns combine with the result filter and search. A filled dot marks filtered columns; **Clear filters** resets all table filters. These filters affect only the display, not summary counts or the exported report.

If no submissions qualify, the status names the selected month and shows the count with unverified eligibility. Check the reporting month and allowed reviewers; an empty reviewer list accepts only blank reviewer cells. Missing Actions settings or evidence leave qualifying rows visible with unresolved results rather than excluding them.

In the Listed names column, **Save list** saves names without refreshing. **Reload TXT** replaces the editor contents with the current file: use it after editing `reviewer_names.txt` in Notepad, then press **Refresh**. The list also loads automatically on startup. Keep the EXE and TXT together in a writable folder. Refreshing saves the reviewer list; unrelated configuration is preserved.

The dashboard uses defaults without a config file. If `config.json` exists beside the executable, it loads those settings; source mode loads it from the working directory. Reviewer edits persist to the TXT. Other dashboard settings apply to the current run and do not change config.json. Existing input/output workbooks are never overwritten. Live captures replace evidence files for the same document and reset attestations; use a separate evidence folder to retain a previously validated snapshot. The dashboard does not monitor Outlook or send email.

Run from source with `python dashboard.py`. To reproduce the executable build:

```powershell
.\.venv\Scripts\python -m pip install -r requirements-build.txt
.\build.ps1
```

The build writes `dist/xCCB_Report.exe` and creates a blank `dist/reviewer_names.txt` only if that file does not already exist. Rebuilding preserves saved reviewer names. PyInstaller's one-file windowed bundle extracts its runtime to a temporary directory when launched. No credentials, source workbook, captured evidence or configuration are bundled.

Captures contain sanitized HTML and JSON manifests, without forms, hidden inputs, scripts, cookies, or authentication headers. Keep them inside approved storage. Up to four document workers overlap I/O; duplicate document references share one fetch per run. Live request starts share a throttle of at least one second between requests, with a 25-second HTTP timeout, a 35-second process timeout, and at most two attempts for transient failures. 401/403 are not retried. A 200 response without the expected iframe/panel is not a successful check. Capture failures and parser warnings appear in the report's Audit sheet. Completion order does not change the saved workbook's manager ordering.

Each output must have a new filename. The source and existing output workbooks are never overwritten. The report retains 43 source columns and their styles/hyperlinks and adds `Flawless submission` as column AR. Header row 3 becomes row 1. Summary and Audit sheets record selection exceptions, evidence, matched identities, and unresolved counts. Audit retains all input row outcomes and their original source row numbers. The main sheet includes qualifying rows even when their Actions classification is unresolved. Missing Assigned By settings or unverified evidence produce a DRAFT with blank classification cells, without suppressing otherwise eligible rows.

The default reporting month is calculated at run time in `Asia/Kuala_Lumpur`; the dashboard can select another month explicitly. Both use an inclusive first day and exclusive first day of the following month. Selecting a historical month does not change the generation timestamp or evidence freshness checks, which still use the real current time. Excel serials respect the source workbook's date epoch. Naive dates/times are treated as local business values; aware timestamps are converted to the business timezone.

## Completeness and classification

Capturing HTML does **not** mark evidence complete. Before a manifest is used, an authorized verifier must check the actual source and fill `scope: "document"`, `validated_by`, and `validation_notes`. Notes must identify the validation of full active/closed/historical listings, pagination, and scope. Set `complete: true` only when the whole relevant history was inspected. The application reparses the captured HTML, verifies its document ID/hash/time, and rejects stale, future, malformed, or mismatched evidence. The default maximum evidence age is 24 hours.

- One matching `Assigned By` identity gives `No` once the evidence scope is confirmed, even if the remaining listing is incomplete.
- No match gives `Yes` only when the full listing is attested complete and parsing has no warnings.
- Otherwise the Yes/No cell stays blank and Audit explains why.
- An empty listing additionally requires an explicitly observed empty-state message: set `empty_state_verified: true` and `empty_state_text` to that exact visible message. An absent/blank table is not sufficient.

Attestations are trusted local inputs, not cryptographic proof. Do not edit a manifest to force a result. A new live capture resets all attestations, so current `--live` reporting remains conservative until that capture has been validated. Automated live `Yes` classification is deliberately not enabled by an unproven global completeness assumption. A verified site-wide extraction contract could remove the per-capture review later; it has not been established in this investigation.

After verifying the manifests in `outputs/evidence`, rerun without `--live` and choose a new output filename to generate the report from that validated evidence.

Evidence is cached by Document ID only for explicitly whole-document scope, within the report run. Failed live capture never falls back to an older successful result. Phase/version/source date remain visible in Audit. Do not reuse this cache for submission-specific results.

## Process flow

```mermaid
flowchart LR
    A[Workbook + configuration] --> B[Filter submission rows]
    B --> C[Qualified rows: resolve document IDs]
    C --> D[Read saved evidence or capture live Actions]
    D --> E[Verify evidence and classify Yes / No / unresolved]
    E --> F[Excel report + Summary + Audit]
    B -->|Excluded or unverified rows| F
```

`__main__.py` parses arguments and starts the report; `workbook.py` loads and writes Excel and coordinates one capture per qualifying document; `rules.py` owns selection and classification; `site.py` fetches and sanitizes HTML; `actions.py` extracts action identities; `evidence.py` creates and verifies manifests. All source rows appear in Audit.

## Validation and findings

```powershell
.\.venv\Scripts\python -m pytest -q
```

See [INVESTIGATION.md](INVESTIGATION.md) for live observations, three-document proof of concept, publishing options and remaining business decisions. Synthetic tests cover rule boundaries, parser failures, verified empty states, evidence integrity, duplicate references, workbook preservation and failed live refresh. UI comparison and the complete production workflow remain unverified.
