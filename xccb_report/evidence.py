"""Local, auditable evidence checkpoints; failures are never cached as success."""

from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path

from bs4 import BeautifulSoup

from .actions import parse_actions, workspace_identity


def write_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(value, indent=2, ensure_ascii=False), encoding="utf-8")
    temp.replace(path)


def inspect_saved(path, url):
    doc_id, canonical = workspace_identity(url)
    path = Path(path).resolve()
    raw = path.read_bytes()
    soup = BeautifulSoup(raw.decode("utf-8-sig"), "html.parser")
    actions, errors = parse_actions(soup, doc_id)
    title = soup.find(id="xccb-document-title")
    return {
        "document_id": doc_id,
        "workspace_url": canonical,
        "document_title": title.get_text(" ", strip=True) if title else "",
        "captured_at": datetime.fromtimestamp(
            path.stat().st_mtime, timezone.utc
        ).isoformat(),
        "source": str(path),
        "sha256": hashlib.sha256(raw).hexdigest(),
        "actions": actions,
        "parse_errors": errors,
        "complete": False,
        "scope": "unconfirmed",
        "validated_by": "",
        "validation_notes": "",
        "kind": "saved-html",
    }


def _read_evidence(path, doc_id, config, now):
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict) or not isinstance(
        value.get("empty_state_text", ""), str
    ):
        raise ValueError("Malformed evidence manifest")
    source = Path(value["source"])
    if not source.is_absolute():
        source = path.parent / source
    raw = source.read_bytes()
    if (
        value["document_id"] != doc_id
        or value["sha256"] != hashlib.sha256(raw).hexdigest()
    ):
        raise ValueError("Evidence source identity or hash mismatch")
    if workspace_identity(value["workspace_url"])[0] != doc_id:
        raise ValueError("Evidence workspace mismatch")
    captured = datetime.fromisoformat(value["captured_at"])
    age = (now - captured).total_seconds() / 3600
    if age < 0 or age > config["evidence_max_age_hours"]:
        raise ValueError("Evidence stale or future-dated")
    return value, raw


def load_document_title(directory, doc_id, config, now):
    try:
        _, raw = _read_evidence(Path(directory) / f"{doc_id}.json", doc_id, config, now)
        title = BeautifulSoup(raw.decode("utf-8-sig"), "html.parser").find(
            id="xccb-document-title"
        )
        return title.get_text(" ", strip=True) if title else ""
    except (OSError, ValueError, KeyError, TypeError):
        return ""


def load_evidence(directory, doc_id, config, now, details=None):
    path = Path(directory) / f"{doc_id}.json"
    if not path.exists():
        return [], False, "No Actions evidence available", ""
    try:
        value, raw = _read_evidence(path, doc_id, config, now)
        soup = BeautifulSoup(raw.decode("utf-8-sig"), "html.parser")
        if details is not None:
            title = soup.find(id="xccb-document-title")
            details["document_title"] = title.get_text(" ", strip=True) if title else ""
        actions, errors = parse_actions(
            soup,
            doc_id,
            (
                value.get("empty_state_text", "")
                if value.get("empty_state_verified") is True
                else ""
            ),
        )
        validated = (
            value.get("scope") == config["action_scope"] == "document"
            and bool(str(value.get("validated_by", "")).strip())
            and bool(str(value.get("validation_notes", "")).strip())
        )
        complete = value.get("complete") is True and validated and not errors
        # Saved evidence alone cannot establish completeness. Human attestation
        # must document all history/active sections and pagination, including zero.
        if not validated:
            return (
                [],
                False,
                "Evidence scope/completeness validation missing",
                str(path),
            )
        return (
            actions,
            complete,
            "; ".join(errors)
            or ("Validated" if complete else "Completeness not attested"),
            str(path),
        )
    except (OSError, ValueError, KeyError, TypeError) as exc:
        return [], False, f"Invalid evidence: {exc}", str(path)
