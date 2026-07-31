from __future__ import annotations

from datetime import datetime, timedelta

from .models import ItineraryEvent


def event_notes(event: ItineraryEvent) -> str:
    lines = [_from_to_line(event)]
    lines.extend(f"{label}: {value}" for label, value in event.notes)
    if event.url:
        lines.append(f"Link: {event.url}")
    return "\n\n".join(lines)


def _from_to_line(event: ItineraryEvent) -> str:
    if event.is_all_day:
        return f"From / To: {event.event_date.isoformat()} all day"

    start = datetime.combine(event.event_date, event.start_time)
    end = datetime.combine(event.event_date, event.end_time)
    if end < start:
        end += timedelta(days=1)
    return f"From / To: {_format_datetime(start)} / {_format_datetime(end)}"


def _format_datetime(value: datetime) -> str:
    return value.strftime("%Y-%m-%d %I:%M %p").replace(" 0", " ")
