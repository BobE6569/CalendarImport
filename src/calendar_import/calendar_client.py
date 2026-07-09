from __future__ import annotations

from datetime import date, datetime, time, timedelta
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from .models import ItineraryEvent


SCOPES = ["https://www.googleapis.com/auth/calendar.events", "https://www.googleapis.com/auth/calendar.readonly"]


class GoogleCalendarClient:
    def __init__(self, project_dir: Path):
        self.project_dir = project_dir
        self.service = build("calendar", "v3", credentials=self._credentials())

    def list_calendars(self) -> list[dict[str, str]]:
        calendars = []
        request = self.service.calendarList().list()
        while request is not None:
            response = request.execute()
            calendars.extend(
                {
                    "id": item["id"],
                    "summary": item.get("summary", item["id"]),
                    "primary": bool(item.get("primary")),
                }
                for item in response.get("items", [])
            )
            request = self.service.calendarList().list_next(request, response)
        return calendars

    def upsert_events(self, calendar_id: str, title_prefix: str, events: list[ItineraryEvent], progress=None) -> dict[str, int]:
        counts = {"created": 0, "updated": 0}
        for index, event in enumerate(events, start=1):
            summary = event.summary(title_prefix)
            existing = self._find_existing_event(calendar_id, summary, event.event_date)
            body = self._event_body(summary, event)
            if existing:
                self.service.events().patch(calendarId=calendar_id, eventId=existing["id"], body=body).execute()
                counts["updated"] += 1
                action = "updated"
            else:
                self.service.events().insert(calendarId=calendar_id, body=body).execute()
                counts["created"] += 1
                action = "created"
            if progress:
                progress(index, len(events), action, summary)
        return counts

    def _find_existing_event(self, calendar_id: str, summary: str, event_date: date) -> dict | None:
        time_min = datetime.combine(event_date - timedelta(days=1), time.min).isoformat() + "Z"
        time_max = datetime.combine(event_date + timedelta(days=2), time.min).isoformat() + "Z"
        response = (
            self.service.events()
            .list(
                calendarId=calendar_id,
                q=summary,
                singleEvents=True,
                orderBy="startTime",
                timeMin=time_min,
                timeMax=time_max,
            )
            .execute()
        )
        candidates = [event for event in response.get("items", []) if event.get("summary") == summary]
        if not candidates:
            return None
        same_date = [event for event in candidates if _event_start_date(event) == event_date.isoformat()]
        return same_date[0] if same_date else candidates[0]

    def _event_body(self, summary: str, event: ItineraryEvent) -> dict:
        body = {
            "summary": summary,
            "location": event.location,
            "description": _description(event),
        }
        if event.url:
            body["source"] = {"title": "Itinerary link", "url": event.url}

        if event.is_all_day:
            end_date = event.event_date + timedelta(days=1)
            body["start"] = {"date": event.event_date.isoformat()}
            body["end"] = {"date": end_date.isoformat()}
        else:
            start = datetime.combine(event.event_date, event.start_time)
            end = datetime.combine(event.event_date, event.end_time)
            if end < start:
                end += timedelta(days=1)
            body["start"] = {"dateTime": start.isoformat(), "timeZone": event.time_zone}
            body["end"] = {"dateTime": end.isoformat(), "timeZone": event.time_zone}
        return body

    def _credentials(self) -> Credentials:
        credentials_path = self.project_dir / "credentials.json"
        token_path = self.project_dir / "token.json"
        credentials = None

        if token_path.exists():
            credentials = Credentials.from_authorized_user_file(str(token_path), SCOPES)
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        if not credentials or not credentials.valid:
            if not credentials_path.exists():
                raise FileNotFoundError(
                    f"Google OAuth credentials were not found at {credentials_path}. "
                    "Download a desktop OAuth credential from Google Cloud and save it there."
                )
            flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), SCOPES)
            credentials = flow.run_local_server(port=0)
        token_path.write_text(credentials.to_json(), encoding="utf-8")
        return credentials


def _description(event: ItineraryEvent) -> str:
    lines = [f"{label}: {value}" for label, value in event.notes]
    if event.url:
        lines.append(f"Link: {event.url}")
    return "\n".join(lines)


def _event_start_date(event: dict) -> str | None:
    start = event.get("start", {})
    if "date" in start:
        return start["date"]
    if "dateTime" in start:
        return start["dateTime"][:10]
    return None
