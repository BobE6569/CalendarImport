from __future__ import annotations

import json
from pathlib import Path


class Preferences:
    def __init__(self, path: Path):
        self.path = path
        self.values = self._load()

    def get(self, key: str, default: str = "") -> str:
        value = self.values.get(key, default)
        return value if isinstance(value, str) else default

    def set(self, key: str, value: str) -> None:
        self.values[key] = value
        self._save()

    def _load(self) -> dict[str, str]:
        if not self.path.exists():
            return {}
        try:
            raw = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return {}
        return raw if isinstance(raw, dict) else {}

    def _save(self) -> None:
        self.path.write_text(json.dumps(self.values, indent=2), encoding="utf-8")

