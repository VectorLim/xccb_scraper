"""Report configuration defaults and validation."""

from datetime import datetime, timezone
import json
import logging
from pathlib import Path

from .reviewers import load_names
from .filtering import month_bounds, normalized
from .organizations import atomic_write, load_settings, locked, user_directory

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


def load_preferences(legacy_names=None):
    settings = load_settings()
    names = settings.get("reviewer_names")
    if names is not None and (
        not isinstance(names, list)
        or any(not isinstance(name, str) or not normalized(name) for name in names)
    ):
        raise ValueError("reviewer_names must be a list of nonblank names")
    filters = settings.get("column_filters", {})
    if not isinstance(filters, dict):
        raise ValueError("column_filters must be an object")
    for column, selection in filters.items():
        if (
            not isinstance(column, str)
            or not isinstance(selection, dict)
            or selection.get("mode", "value") not in {"value", "month", "day"}
            or not isinstance(selection.get("values"), list)
            or any(not isinstance(value, str) for value in selection["values"])
            or not isinstance(selection.get("all", False), bool)
        ):
            raise ValueError("Column filters require a mode and a list of text values")
    if names is None and legacy_names is not None and Path(legacy_names).exists():
        settings["reviewer_names"] = load_names(legacy_names)
        save_preferences({"reviewer_names": settings["reviewer_names"]})
    return settings


def save_preferences(updates):
    path = user_directory() / "settings.json"
    with locked(path):
        settings = load_settings(path)
        settings.update(updates)
        atomic_write(path, lambda stream: json.dump(settings, stream, indent=2))


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
    preferences = load_preferences(Path(path).resolve().parent / "reviewer_names.txt")
    if "reviewer_names" in preferences:
        config["reviewer_names"] = config["assigned_by_names"] = preferences["reviewer_names"]
    return config
