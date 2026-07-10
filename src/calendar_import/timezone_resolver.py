from __future__ import annotations

from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError


DEFAULT_TIME_ZONE = "America/New_York"

BUILT_IN_LOCATION_TIME_ZONES = {
    "atlanta": "America/New_York",
    "boston": "America/New_York",
    "chicago": "America/Chicago",
    "denver": "America/Denver",
    "florence": "Europe/Rome",
    "firenze": "Europe/Rome",
    "las vegas": "America/Los_Angeles",
    "london": "Europe/London",
    "los angeles": "America/Los_Angeles",
    "milan": "Europe/Rome",
    "milano": "Europe/Rome",
    "naples": "Europe/Rome",
    "napoli": "Europe/Rome",
    "new york": "America/New_York",
    "orlando": "America/New_York",
    "paris": "Europe/Paris",
    "philadelphia": "America/New_York",
    "rome": "Europe/Rome",
    "roma": "Europe/Rome",
    "san francisco": "America/Los_Angeles",
    "venice": "Europe/Rome",
    "venezia": "Europe/Rome",
    "washington dc": "America/New_York",
    "washington, dc": "America/New_York",
}


class TimeZoneResolver:
    def __init__(self, mapping_path: Path, prompt_for_time_zone=None):
        self.mapping_path = mapping_path
        self.prompt_for_time_zone = prompt_for_time_zone
        self.location_names: dict[str, str] = {}
        self.custom_map = self._load_custom_map()

    def resolve(self, location: str, previous_time_zone: str | None = None) -> str:
        value = (location or "").strip()
        if not value:
            return previous_time_zone or DEFAULT_TIME_ZONE

        if self._is_valid_time_zone(value):
            return value

        normalized = self._normalize(value)
        custom = self.custom_map.get(normalized)
        if custom and self._is_valid_time_zone(custom):
            return custom

        built_in = BUILT_IN_LOCATION_TIME_ZONES.get(normalized)
        if built_in:
            return built_in

        if self.prompt_for_time_zone:
            prompted = self.prompt_for_time_zone(value, previous_time_zone or DEFAULT_TIME_ZONE)
            if prompted and self._is_valid_time_zone(prompted):
                self.custom_map[normalized] = prompted
                self.location_names[normalized] = value
                self._save_custom_map()
                return prompted

        return previous_time_zone or DEFAULT_TIME_ZONE

    def _load_custom_map(self) -> dict[str, str]:
        if not self.mapping_path.exists():
            return {}
        try:
            lines = self.mapping_path.read_text(encoding="utf-8").splitlines()
        except OSError:
            return {}

        output = {}
        for line in lines:
            cleaned = line.strip()
            if not cleaned or cleaned.startswith("#") or "=" not in cleaned:
                continue
            location, time_zone = cleaned.split("=", 1)
            if location.strip() and time_zone.strip():
                normalized = self._normalize(location)
                self.location_names[normalized] = location.strip()
                output[normalized] = time_zone.strip()
        return output

    def _save_custom_map(self) -> None:
        lines = [
            "# CalendarImport location time-zone mappings",
            "# Format: Location=IANA time zone",
            "# Example: Rome=Europe/Rome",
            "",
        ]
        lines.extend(
            f"{self.location_names.get(location, location)}={time_zone}"
            for location, time_zone in sorted(self.custom_map.items())
        )
        self.mapping_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    @staticmethod
    def _normalize(value: str) -> str:
        return " ".join(value.lower().replace("\u00a0", " ").split())

    @staticmethod
    def _is_valid_time_zone(value: str) -> bool:
        try:
            ZoneInfo(value)
        except ZoneInfoNotFoundError:
            return False
        return True


def mapping_path_for_workbook(project_dir: Path, workbook_path: Path) -> Path:
    safe_name = "".join(
        character if character.isalnum() or character in ("-", "_") else "_"
        for character in workbook_path.stem.strip()
    ).strip("_")
    if not safe_name:
        safe_name = "workbook"
    return project_dir / f"{safe_name}_timezones.txt"
