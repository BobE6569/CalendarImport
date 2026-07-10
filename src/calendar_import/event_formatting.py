from __future__ import annotations

from .models import ItineraryEvent


def event_notes(event: ItineraryEvent) -> str:
    lines = [f"{label}: {value}" for label, value in event.notes]
    if event.url:
        lines.append(f"Link: {event.url}")
    return "\n\n".join(lines)

