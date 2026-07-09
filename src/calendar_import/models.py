from __future__ import annotations

from dataclasses import dataclass
from datetime import date, time


@dataclass(frozen=True)
class ItineraryEvent:
    row_number: int
    event_date: date
    location: str
    start_time: time | None
    end_time: time | None
    url: str
    notes: list[tuple[str, str]]
    time_zone: str

    @property
    def is_all_day(self) -> bool:
        return self.start_time is None and self.end_time is None

    def summary(self, title_prefix: str) -> str:
        clean_title = title_prefix.strip()
        clean_location = self.location.strip()
        if clean_location:
            return f"{clean_title} - {clean_location}"
        return clean_title

