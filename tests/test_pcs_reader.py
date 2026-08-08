from datetime import date, time

from openpyxl import Workbook

from calendar_import.pcs_reader import PCS_EVENT_LOCATION, read_pcs_club_events


def test_pcs_reader_creates_reservation_and_event_items(tmp_path):
    path = tmp_path / "pcs.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Date", "Event", "Location", "From", "To", "Reservations Open"])
    sheet.append([date(2026, 9, 12), "Wine Dinner", "Main Dining Room", time(18, 0), "", date(2026, 8, 1)])
    workbook.save(path)

    events = read_pcs_club_events(path)

    assert len(events) == 2
    reservation, event = events
    assert reservation.summary("") == "PCS Reservation: Wine Dinner"
    assert reservation.event_date == date(2026, 8, 1)
    assert reservation.start_time == time(9, 0)
    assert reservation.end_time == time(9, 0)
    assert reservation.location == ""
    assert reservation.note_override == "2026-09-12 6:00 PM 10:00 PM Wine Dinner\nMain Dining Room"
    assert reservation.reminder_minutes_before_start == 0

    assert event.summary("") == "PCS: Wine Dinner"
    assert event.event_date == date(2026, 9, 12)
    assert event.start_time == time(18, 0)
    assert event.end_time == time(22, 0)
    assert event.location == PCS_EVENT_LOCATION
    assert event.note_override == "2026-09-12 6:00 PM 10:00 PM Wine Dinner\nMain Dining Room"
    assert event.reminder_minutes_before_start == 60
