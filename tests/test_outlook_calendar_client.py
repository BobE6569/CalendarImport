from datetime import date, datetime

from calendar_import.event_formatting import event_notes
from calendar_import.models import ItineraryEvent
from calendar_import.outlook_calendar_client import IANA_TO_WINDOWS_TIME_ZONE, _outlook_datetime


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

    assert event_notes(event) == "Details: Tour\n\nConfirmation: ABC123\n\nLink: https://example.com"


def test_outlook_datetime_filter_format():
    assert _outlook_datetime(datetime(2026, 8, 1, 9, 30)) == "08/01/2026 09:30 AM"


def test_outlook_time_zone_map_includes_buenos_aires():
    assert IANA_TO_WINDOWS_TIME_ZONE["America/Buenos_Aires"] == "Argentina Standard Time"
    assert IANA_TO_WINDOWS_TIME_ZONE["America/Argentina/Buenos_Aires"] == "Argentina Standard Time"


def test_outlook_time_zone_map_includes_santiago_and_lima():
    assert IANA_TO_WINDOWS_TIME_ZONE["America/Santiago"] == "Pacific SA Standard Time"
    assert IANA_TO_WINDOWS_TIME_ZONE["America/Lima"] == "SA Pacific Standard Time"
