"""Windows-integrated, GET-only access to the observed xCCB HTML pages."""
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import re
import subprocess
from threading import Lock
import time
from urllib.parse import parse_qs, urljoin, urlparse

from bs4 import BeautifulSoup, Comment

from .actions import workspace_identity
from .evidence import inspect_saved, write_json


PS_GET = r'''
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
try {
  $r = Invoke-WebRequest -Uri $env:XCCB_READ_URL -UseDefaultCredentials -UseBasicParsing -TimeoutSec 25 -MaximumRedirection 0
  @{status=[int]$r.StatusCode; url=$r.BaseResponse.ResponseUri.AbsoluteUri; html=$r.Content} | ConvertTo-Json -Compress
} catch {
  $status = 0
  if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
  @{status=$status; url=$env:XCCB_READ_URL; html=''} | ConvertTo-Json -Compress
}
'''


class WindowsReader:
    def __init__(self, interval=1.0, attempts=2, cancel=None):
        self.interval, self.attempts, self.last = max(1.0, interval), attempts, 0.0
        self.lock = Lock()
        self.cancel = cancel

    def check_cancel(self):
        if self.cancel is not None and self.cancel.is_set():
            raise OSError("Stopped by user")

    def get(self, url):
        if os.name != "nt":
            raise OSError("Windows Integrated Authentication requires Windows")
        parsed = urlparse(url)
        if parsed.scheme != "https" or parsed.netloc != "xccb2.intel.com":
            raise ValueError("Only the approved HTTPS xCCB host is permitted")
        for attempt in range(self.attempts):
            self.check_cancel()
            with self.lock:
                time.sleep(max(0, self.interval - (time.monotonic() - self.last)))
                self.last = time.monotonic()
            self.check_cancel()
            env = dict(os.environ, XCCB_READ_URL=url)
            try:
                result = subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", PS_GET],
                                        env=env, capture_output=True, timeout=35, encoding="utf-8",
                                        creationflags=subprocess.CREATE_NO_WINDOW)
                if result.returncode:
                    raise OSError("Windows HTTP reader failed")
                response = json.loads(result.stdout.lstrip("\ufeff"))
                status = response["status"]
            except (subprocess.TimeoutExpired, json.JSONDecodeError):
                status = 0
            if status == 200:
                final = urlparse(response["url"])
                if final.scheme != "https" or final.netloc != "xccb2.intel.com":
                    raise OSError("Unexpected authentication redirect")
                return response["html"]
            if status in {401, 403}:
                raise OSError(f"Windows authentication/access failed (HTTP {status})")
            if status not in {0, 429, 502, 503, 504} or attempt == self.attempts - 1:
                raise OSError(f"xCCB GET failed (HTTP {status})")
            time.sleep(2 ** attempt)


def sanitized(node):
    for hidden in node.find_all(["script", "style", "input", "head"]):
        hidden.decompose()
    for comment in node.find_all(string=lambda value: isinstance(value, Comment)):
        comment.extract()
    for form in node.find_all("form"):
        form.unwrap()
    for tag in [node, *node.find_all(True)]:
        tag.attrs = {k: v for k, v in tag.attrs.items() if k in {"id", "class", "href", "title", "rowspan", "colspan"}}
    return str(node)


def capture(reader, url, directory):
    doc_id, canonical = workspace_identity(url)
    workspace = BeautifulSoup(reader.get(canonical), "html.parser")
    frames = workspace.find_all("iframe", id="iFrameActions")
    if len(frames) != 1:
        page = " ".join(workspace.get_text(" ").split())
        if "Document not found" in page:
            deleted = re.search(r"deleted by .+? on [0-9/]+ [0-9:]+ [AP]M", page)
            raise ValueError(f"Document deleted on xCCB ({deleted.group(0)})" if deleted else "Document not found on xCCB")
        raise ValueError("Workspace Actions iframe missing; possible authentication or loading failure")
    frame_url = urljoin(canonical, frames[0].attrs.get("src", ""))
    parsed = urlparse(frame_url)
    query = parse_qs(parsed.query)
    if (parsed.scheme != "https" or parsed.netloc != "xccb2.intel.com"
            or parsed.path.casefold() != "/documents/documentcomponents.aspx"
            or query.get("Component") != ["Actions"] or query.get("DocumentID") != [doc_id]):
        raise ValueError("Unexpected Actions iframe route or document identity")
    # Complete HTTP response body is received before parsing; no browser timing heuristic.
    frame = BeautifulSoup(reader.get(frame_url), "html.parser")
    clean = sanitized(frame.body or frame)
    directory = Path(directory).resolve()
    directory.mkdir(parents=True, exist_ok=True)
    source = directory / f"{doc_id}.html"
    source.write_text(clean, encoding="utf-8")
    result = inspect_saved(source, canonical)
    # A full live GET of the document's own Actions iframe is the whole-document listing.
    result |= {"frame_url": frame_url, "captured_at": datetime.now(timezone.utc).isoformat(), "kind": "live-html-get",
               "scope": "document", "complete": not result["parse_errors"],
               "validated_by": "Live xCCB capture", "validation_notes": f"GET {frame_url}"}
    write_json(directory / f"{doc_id}.json", result)
    return result
