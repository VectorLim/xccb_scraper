# Investigation — 30 October 2026

> Updated requirements: Author Manager is a grouping field, not an organization filter. Phase accepts all values except those containing WLA (case-insensitive), including blank phases. Reviewer(s) accepts the editable allowed name list plus blank cells. The historical candidate-list/mapping decisions and old dry-run selection counts below are superseded. See README.md for the current dashboard, executable build, filtering and evidence behavior. Actions Assigned By matches a separate target list, not the Excel Reviewer(s) or Author field; Outlook automation is not implemented.

Reviewer-list update: the dashboard has scrollable name editors, Save list, Reload TXT and Update & run controls. Allowed reviewers persist in reviewer_names.txt beside the executable and reload on startup. Update reruns the full report using the edited list and creates a new workbook after each successful run. No specific allowed reviewers have been supplied; the initial list is empty, so only blank Reviewer(s) cells pass until names are entered. The earlier selection counts below describe prior filter versions, not the current empty-list population.

Final reviewer-list verification: 52 offline tests passed both in the main run and an independent test run. Coverage includes listed names plus blanks, nonmatching and similar names, TXT save/deduplication/BOM/reload, dashboard reruns with changed lists and unique outputs, and preservation of an existing workbook at the former temporary filename. The executable was rebuilt with this implementation. Packaged startup was checked during this update; desktop visual inspection and live Actions/report execution remain unverified.

Previous phase-only update: 47 offline tests passed and the executable was rebuilt. At that point, with the former nonblank-reviewer rule, the supplied October workbook had 102 qualified, 1,201 excluded and one unverified row. Those counts predate the current allowed-list-plus-blanks rule.

## Updated implementation verification

The Tkinter dashboard and `dist/xCCB_Report.exe` implement the revised filters and manager summaries. The 38-test offline suite passed, including background dashboard generation, output overwrite protection and the live-capture timestamp regression. An independent test runner also passed all 38 tests. Offline processing of the supplied workbook for October 2026 produced 97 qualifying rows across 55 manager groups, 1,206 excluded rows and one unverified row. All 1,304 rows remain in Audit. With no validated Actions or target list supplied, all 97 classifications remain unresolved and the workbook is DRAFT. The packaged executable started with the expected responsive dashboard window; desktop automation was unavailable, so visual inspection and report generation through the packaged UI were not performed. No live requests or email sends were made during this update.

## Evidence and access

Authoritative references are the project-local workbook and saved HTML/companion assets. The workbook SHA-256 is `883564e7f53aaad4d804f9926ed25b82b2fc0bb077d3cedafa8a71f7893db98d`; it matched the Downloads copy inspected initially. The project was initially empty and the supplied references appeared there during the session. No original files were edited.

The workbook has one sheet, 43 columns A:AQ, headers on row 3, 1,304 data rows, 474 distinct document references, 125 reviewer names and 163 manager names. All 1,304 workspace cells contain hyperlink targets. There are no embedded configuration sheets, comments, defined names or organization mappings. The profile artifact records candidate names without treating them as an approved list.

Browser integration reported `The browser access host is unavailable`. Native Edge observation then stopped because Computer Use could not verify its current URL; no further UI actions were attempted. Thus **manual browser confirmation was not performed**. Read-only PowerShell `Invoke-WebRequest -UseDefaultCredentials` succeeded against the live site using the running Windows account. No password, token, cookie export or security-setting changes were needed.

## Actual Actions implementation

The live workspace for document 2369358 returned HTTP 200 and an iframe:

```html
<iframe id="iFrameActions" src="DocumentComponents.aspx?Component=Actions&DocumentID=2369358" ...>
```

Fetching that observed same-host URL with a normal authenticated GET returned the server-rendered Actions table. This is an HTML component, not evidence of a JSON/API service. The pipeline always reads the workspace and resolves its real iframe source; it does not construct an assumed component endpoint.

Each action has a linked `CompleteDocumentAction.aspx?ActionID=...&DocumentID=...` record and an inline `Assigned By:` span. An adjacent cell lists recipients; that cell is not used for classification. Actions use paired table rows with rowspans, where the companion row shows status. Individual action records need not be opened to extract assigners. All form submission, edit and completion controls are left untouched.

The saved companion `DocumentComponents(4).html` actually exists and contains 19 actions, contrary to the attachment's preliminary assumption that it was missing. It points to the same component URL, but the selection of the extraction mechanism is also supported by live requests.

Live sanitized captures and manifests are under `outputs/live-evidence/`:

| Document ID | Actions | Assigned By breakdown |
|---|---:|---|
| 2369358 | 19 | Business Rule (System assigned): 17; Kang, Wei Yu: 1; Jiang, Wei3: 1 |
| 2246498 | 10 | Business Rule: 6; Rajan, Divya: 1; Beverlin, Lucas P: 2; Alzola, Joaquin: 1 |
| 2250366 | 4 | Business Rule: 2; Rajan, Divya: 1; Mathesen, Logan: 1 |

No parser warnings or visible pagination controls were found in these captured tables. A broad raw-text `pager` search matched surrounding page/script content, which is why extraction checks actual controls instead of that substring. Absence of a visible pager does not establish that all historical phases/versions are represented. No empty live document was validated. The saved workspace says Final, version 219, and an approval number containing 2021, while the saved Actions dates are October 2026; this is insufficient to establish submission-level scope. All capture manifests remain incomplete and unconfirmed.

The saved workflow links ATM P CCB Admin Review to `ForumDetails.aspx?ForumID=1001340`. A live GET returned HTTP 200 with **Access Denied** in its page, so HTTP success alone does not imply authorization. A reviewer membership roster was not retrieved. Metadata lists affected sites including KuAT and a document-selected department manager; neither establishes a general manager-to-organization mapping.

An independent Luna helper fetched the same three live component URLs with Windows credentials and a separate regex check. All ordered Action IDs and inline Assigned By identities matched the manifests (19/10/4); the observed Closed/Open counts were 19/0, 10/0 and 4/0, and no pager controls were found. Its strict span-only regex missed nested markup in some records, so it independently compared the labeled inline text segments instead. This confirms the current HTTP responses, not historical completeness or browser/UI rendering. A final GET-only run of document 2369358 after disabling automatic redirects again retrieved 19 Actions without parser warnings.

## Publishing investigation

The linked portal responded with a Report Server web portal shell. Its authenticated standard REST catalog at `/reports/api/v2.0/CatalogItems` returned 24 visible items, including `/xCCB Reports/ForumReviewers (Includes Metadata)` (report ID `6f69e537-6206-4ca5-b907-0f7a1150b9d4`). This confirms that an existing source report can be reused.

Read-only capability checks on that report showed Execute and View, Read Properties, Create/Read/Update/Delete Subscription, List Report History and Comment on Reports. No subscription was created or changed; the subscription list returned zero visible entries. The server's `/Extensions` collection lists `Report Server Email`, `Report Server FileShare`, `EXCEL`, and `EXCELOPENXML`. These prove registered server capabilities, not SMTP configuration, successful delivery, folder upload rights or authorization to edit the existing report.

| Option | Practical path | Setup still required |
|---|---|---|
| Internal dashboard / report portal | Run the Python job under an approved Windows identity; expose its results through an approved data source and a new paginated report or approved dashboard | Report/data-source owner; permission to create a custom report; approved storage; refresh identity and cadence; access policy; confirmed business rules |
| Scheduled Excel email | Prefer the portal's existing subscription infrastructure once a report consuming the computed results exists; otherwise an approved scheduler/mail service must send the Python workbook | Confirm recipients, schedule, SMTP/mail policy, service account and destination; approve creation of subscription or mail job |

An existing report subscription renders **that report's dataset**. It will not automatically include Python's extra column or mail an arbitrary local workbook. The computed data must first be integrated into an approved report/data source, or delivered by an approved separate job. No built-in Python execution or arbitrary extra computed field was verified. Start with the local workbook, then reuse the report portal if its owner can host the augmented dataset.

Microsoft documents the [Reporting Services REST API](https://learn.microsoft.com/en-us/sql/reporting-services/developer/rest-api) and [web portal subscriptions with Excel delivery](https://learn.microsoft.com/en-us/sql/reporting-services/working-with-subscriptions-web-portal). Those product capabilities are consistent with the live registered extensions; local delivery has not been tested.

## Remaining decisions and validation

1. Supply the intended reviewer list and the separate Assigned By list. The nine admin-associated reviewers in the export are candidates, not proof of the requested lists.
2. Supply an approved mapping from Author Manager identities to KuAT AE/TE, or explicitly approve a broader KuAT mapping.
3. Confirm whether WLA CCB Final / WLA CCB Preliminary and PrePilot are included. Exact plain Final/Preliminary rules are the only defaults.
4. Confirm whole-document versus submission/phase/version Actions scope, and verify live history/pagination and empty states before enabling Yes results.
5. Confirm recipients/destination, job identity, schedule, and portal owner before publishing.

The dry-run workbook, `outputs/2026-10-30/Flawless_Submission_DRAFT_verified.xlsx`, is clearly marked DRAFT. With unconfigured lists/mapping and the explicit plain phase defaults, 1,188 rows fail known filters and 116 remain unresolved for eligibility. There are no selected rows and no Yes/No claims. Audit records all 1,304 source rows, with working workspace hyperlinks. WLA variants remain excluded by the documented default, subject to the pending business decision. The 22-test suite passed after correcting cross-workbook cell and column style copying to register actual style components rather than source-workbook indices; literal formula-like source text also retains its original Excel cell type. The saved workbook reopened successfully with openpyxl, with zero error cells, 44 report columns and 1,304 audit hyperlinks. Artifact Tool imported and inspected it and produced previews of all three sheets, which were visually checked; its Node process returned exit 1 without a diagnostic after writing the images, so that renderer invocation is not reported as a clean check. Native Excel rendering, manual UI comparison and a fully classified real-data end-to-end run remain outstanding.
