from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict

DEFAULT_OPTIONS: Dict[str, Any] = {
    "diagnoses": [],
    "medications": [],
    "standard_phrases": [],
    "investigations": [],
    "symptoms_anamnesis": [],
    "meta": {"note": "Liste pentru ajutor de redactare, nu decizie clinică."},
}


def load_options(path: str | Path) -> Dict[str, Any]:
    path = Path(path)
    if not path.exists():
        return DEFAULT_OPTIONS.copy()
    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
    except json.JSONDecodeError:
        return DEFAULT_OPTIONS.copy()
    merged = DEFAULT_OPTIONS.copy()
    merged.update(data)
    return merged


def save_options(path: str | Path, data: Dict[str, Any]) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def labels(items: list[dict[str, Any]], include_counts: bool = False) -> list[str]:
    out = []
    for item in items:
        label = str(item.get("label") or item.get("text") or "").strip()
        count = item.get("count")
        if not label:
            continue
        if include_counts and count is not None:
            out.append(f"{label}  ·  n={count}")
        else:
            out.append(label)
    return out


def strip_count(label: str) -> str:
    return label.split("  ·  n=")[0].strip()


def add_unique_item(options: Dict[str, Any], key: str, value: str, label_key: str = "label") -> bool:
    value = value.strip()
    if not value:
        return False
    current = options.setdefault(key, [])
    normalized = value.casefold()
    for item in current:
        existing = str(item.get(label_key) or item.get("text") or "").strip().casefold()
        if existing == normalized:
            return False
    current.append({label_key: value})
    return True
