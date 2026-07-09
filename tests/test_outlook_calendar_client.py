from datetime import date, datetime

from calendar_import.models import ItineraryEvent
from calendar_import.outlook_calendar_client import _body, _outlook_datetime


def test_outlook_body_includes_notes_and_link():
    event = ItineraryEvent(
        row_number=2,
        event_date=date(2026, 8, 1),
        location="Rome",
        start_time=None,
        end_time=None,
        url="https://example.com",
        notes=[("Details", "Tour"), ("Confirmation", "ABC123")],
        time_zone="Europe/Rome",
    )

    assert _body(event) == "Details: Tour\nConfirmation: ABC123\nLink: https://example.com"


def test_outlook_datetime_filter_format():
    assert _outlook_datetime(datetime(2026, 8, 1, 9, 30)) == "08/01/2026 09:30 AM"
