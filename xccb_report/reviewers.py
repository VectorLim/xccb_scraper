"""Editable reviewer-list persistence, shared by the CLI and dashboard."""
from pathlib import Path

from .rules import normalized


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


def save_names(path, text):
    path = Path(path)
    names = parse_names(text)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text("\n".join(names) + ("\n" if names else ""), encoding="utf-8")
    temp.replace(path)
    return names
