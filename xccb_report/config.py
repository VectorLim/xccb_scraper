"""Report configuration defaults and validation."""

from datetime import datetime, timezone
import json
import logging
from pathlib import Path

from .reviewers import load_names
from .rules import month_bounds, normalized

DEFAULTS = {
    "timezone": "Asia/Kuala_Lumpur",
    "assigned_by_names": [],
    "reviewer_names": [],
    "action_scope": "unconfirmed",
    "evidence_max_age_hours": 24,
}
LEGACY_FILTER_KEYS = {
    "reviewer_separator",
    "manager_organizations",
    "allowed_organizations",
    "phases",
}


def load_config(path):
    supplied = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    if (
        not isinstance(supplied, dict)
        or set(supplied) - set(DEFAULTS) - LEGACY_FILTER_KEYS
    ):
        raise ValueError("Configuration must be an object with supported keys")
    for key in LEGACY_FILTER_KEYS & supplied.keys():
        supplied.pop(key)
        logging.warning("Ignoring obsolete filter setting: %s", key)
    config = DEFAULTS | supplied
    for key in ("assigned_by_names", "reviewer_names"):
        if not isinstance(config[key], list) or any(
            not isinstance(v, str) or not normalized(v) for v in config[key]
        ):
            raise ValueError(f"{key} must be a list of nonblank names")
    if config["action_scope"] not in {"unconfirmed", "document"}:
        raise ValueError(
            "Only whole-document scope is supported; submission/phase scope requires verified metadata"
        )
    if (
        not isinstance(config["evidence_max_age_hours"], (int, float))
        or config["evidence_max_age_hours"] <= 0
    ):
        raise ValueError("Evidence age must be positive")
    month_bounds(datetime.now(timezone.utc), config["timezone"])
    reviewer_path = Path(path).resolve().parent / "reviewer_names.txt"
    if reviewer_path.exists():
        # One listed-names file drives both the Reviewer(s) filter and the Assigned by check.
        config["reviewer_names"] = config["assigned_by_names"] = load_names(
            reviewer_path
        )
    return config
