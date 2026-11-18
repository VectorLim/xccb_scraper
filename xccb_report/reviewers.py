"""Editable reviewer-list persistence, shared by the CLI and dashboard."""

from pathlib import Path

from .filtering import normalized


def parse_names(text):
    names, seen = [], set()
    for line in text.splitlines():
        name = " ".join(line.split())
        key = normalized(name)
        if key and key not in seen:
            names.append(name)
            seen.add(key)
    return names


def load_names(path):
    return parse_names(Path(path).read_text(encoding="utf-8-sig"))
