# Deploy xCCB Report on the intranet

The Linux container serves the browser UI and temporary report jobs. The Windows worker uses its signed-in corporate account for xCCB and the existing DataSyncX/Oracle setup for HR lookup. Browsers on the internal network use the app without individual sign-in.

## Set up the Docker host

1. Check out a release on the internal Docker host. Create a random worker token in an approved secrets location and install a separate file with the same token on the worker PC. For example, on Linux, `openssl rand -hex 32` produces a 64-character token. Restrict access to the file. For the Compose file-backed secret, the mounted file must be readable by container UID 10001. On a standard Linux Docker host, assigning the file to that UID with mode `0400` meets this requirement.
2. Set `XCCB_WORKER_SECRET_HOST_FILE` to the host token file path. Optionally set `XCCB_REPORT_CONFIG_FILE` to a readable report configuration file. The default is `config.example.json`. These paths may be supplied in the deployment checkout's ignored `.env` file. Compose binds to `127.0.0.1:8080` by default.
3. Configure the host's internal HTTPS proxy to forward the site to that loopback port. Restrict the site to the intended corporate network and restrict `/worker/` to the worker network or machine. Forward the worker's `Authorization` header. Support request bodies up to the larger of the upload and generated-report limits and avoid caching API responses. The worker can use your approved CA bundle through `REQUESTS_CA_BUNDLE` when required.
4. Run `docker compose up -d --build --wait`. `/healthz` checks the web process. `/api/status` shows whether reporting is enabled, the worker is online, and how many jobs are queued or running.

## Start the Windows worker

Use the same release checkout and an approved Python 3.11–3.13 environment with the report requirements, `defusedxml==0.7.1`, DataSyncX 1.1.6, and its existing Oracle client/TNS setup. Run this in PowerShell under the account that can open live xCCB workspaces.

```powershell
.\.venv-build\Scripts\python.exe -m pip install defusedxml==0.7.1
$env:XCCB_WEB_URL = 'https://xccb-report.example.internal'
$env:XCCB_WORKER_SECRET_FILE = 'C:\Secure\xccb-worker-token'
.\.venv-build\Scripts\python.exe -m xccb_report.web_worker
```

Run one worker. Keep the PC signed in and connected to the corporate network. The worker makes outbound HTTPS calls and uses the token only for the report app. The xCCB requests use Windows Integrated Authentication. For automatic startup, run this command at that user's Windows logon. Restart the worker after updating its checkout or token.

## Verify and enable

- Confirm with the xCCB and HR data owners that the full document scope available to the worker account and the workbook's Audit fields, including employee WWIDs, may be served to the intended corporate audience without individual sign-in.
- Compare a representative live report with the desktop report. Confirm the portal's active, closed, historical and paginated Actions listings before accepting hosted `Yes` classifications. Confirm xCCB and Oracle access using the worker's actual Windows session.
- Set `XCCB_ENABLE_REPORTS=1` in the deployment environment and recreate the web service with the Compose command above. Run an approved smoke workbook through the browser and download its Excel report.

## Redeploy or roll back

Check `/api/status` for active jobs before replacing the web container. Update the deployment checkout to the new release and run:

```sh
docker compose up -d --build --wait
```

To roll back, select the previous release in the deployment checkout and run the same command. If a release changes the worker, update and restart it from the matching checkout. Use `docker compose logs web` and the Windows worker console to diagnose failures.

The service runs one report at a time. A container restart clears job state, so interrupted jobs need a fresh upload. Reports expire after one hour or when the same browser session submits a replacement. Workbooks, evidence and organization lookups use temporary job storage. The worker removes its scratch folder on normal exit from a job. After a hard PC shutdown, remove any abandoned `xccb-worker-*` folders from the worker account's temporary directory once the worker is stopped. Browser preferences survive tab reloads and clear on tab closure.

Defaults are 25 MiB per uploaded workbook, 50 MiB per generated report, 250 MiB of expanded archive contents, 12 retained jobs, and a one-hour queue and run deadline. Set `XCCB_MAX_UPLOAD_MIB`, `XCCB_MAX_REPORT_MIB`, `XCCB_MAX_EXPANDED_MIB`, `XCCB_MAX_JOBS`, or `XCCB_JOB_TIMEOUT_SECONDS` on the Docker host if representative reports require different limits.
