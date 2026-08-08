from __future__ import annotations

from datetime import date, datetime, time, timedelta

from .event_formatting import event_notes
from .models import ItineraryEvent


OL_APPOINTMENT_ITEM = 1
OL_FOLDER_CALENDAR = 9
OL_MEETING = 1
OL_REQUIRED = 1

IANA_TO_WINDOWS_TIME_ZONE = {
    "America/Argentina/Buenos_Aires": "Argentina Standard Time",
    "America/Buenos_Aires": "Argentina Standard Time",
    "America/Chicago": "Central Standard Time",
    "America/Denver": "Mountain Standard Time",
    "America/Lima": "SA Pacific Standard Time",
    "America/Los_Angeles": "Pacific Standard Time",
    "America/New_York": "Eastern Standard Time",
    "America/Santiago": "Pacific SA Standard Time",
    "Europe/London": "GMT Standard Time",
    "Europe/Paris": "Romance Standard Time",
    "Europe/Rome": "W. Europe Standard Time",
}


class OutlookCalendarClient:
    def __init__(self):
        try:
            import win32com.client
        except ImportError as exc:
            raise ImportError(
                "Outlook import requires pywin32. Run the setup script or install requirements.txt."
            ) from exc

        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")

    def list_calendars(self) -> list[dict[str, str]]:
        calendars = []
        default_calendar = self.namespace.GetDefaultFolder(OL_FOLDER_CALENDAR)
        default_entry_id = default_calendar.EntryID

        for store_index in range(1, self.namespace.Folders.Count + 1):
            root = self.namespace.Folders.Item(store_index)
            calendars.extend(self._calendar_folders(root, root.Name, default_entry_id))

        calendars.sort(key=lambda item: (not item["primary"], item["summary"].lower()))
        return calendars

    def upsert_events(
        self,
        calendar_id: str,
        title_prefix: str,
        events: list[ItineraryEvent],
        invitee: str = "",
        progress=None,
    ) -> dict[str, int]:
        folder = self._folder_from_id(calendar_id)
        counts = {"created": 0, "updated": 0}

        for index, event in enumerate(events, start=1):
            subject = event.summary(title_prefix)
            appointment = self._find_existing_event(folder, subject, event)
            if appointment:
                counts["updated"] += 1
                action = "updated"
            else:
                appointment = folder.Items.Add(OL_APPOINTMENT_ITEM)
                counts["created"] += 1
                action = "created"

            self._apply_event(appointment, subject, event, invitee)
            appointment.Save()

            if progress:
                progress(index, len(events), action, subject)

        return counts

    def _calendar_folders(self, folder, prefix: str, default_entry_id: str) -> list[dict[str, str]]:
        calendars = []
        try:
            is_calendar = folder.DefaultItemType == OL_APPOINTMENT_ITEM
        except Exception:
            is_calendar = False

        if is_calendar:
            calendars.append(
                {
                    "id": f"{folder.StoreID}|{folder.EntryID}",
                    "summary": prefix,
                    "primary": folder.EntryID == default_entry_id,
                }
            )

        try:
            child_count = folder.Folders.Count
        except Exception:
            return calendars

        for child_index in range(1, child_count + 1):
            child = folder.Folders.Item(child_index)
            calendars.extend(self._calendar_folders(child, f"{prefix} / {child.Name}", default_entry_id))

        return calendars

    def _folder_from_id(self, calendar_id: str):
        store_id, entry_id = calendar_id.split("|", 1)
        return self.namespace.GetFolderFromID(entry_id, store_id)

    def _find_existing_event(self, folder, subject: str, event: ItineraryEvent):
        start_window = datetime.combine(event.event_date - timedelta(days=1), time.min)
        end_window = datetime.combine(event.event_date + timedelta(days=2), time.min)

        items = folder.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")
        restriction = (
            f"[Start] >= '{_outlook_datetime(start_window)}' "
            f"AND [Start] < '{_outlook_datetime(end_window)}'"
        )

        try:
            restricted = items.Restrict(restriction)
        except Exception:
            restricted = items

        for appointment in restricted:
            try:
                if appointment.Subject == subject and appointment.Start.date() == event.event_date:
                    return appointment
            except Exception:
                continue

        return None

    def _apply_event(self, appointment, subject: str, event: ItineraryEvent, invitee: str = "") -> None:
        appointment.Subject = subject
        appointment.Location = event.location
        appointment.Body = event_notes(event)
        appointment.AllDayEvent = bool(event.is_all_day)
        self._apply_reminder(appointment, event.reminder_minutes_before_start)

        if event.is_all_day:
            appointment.Start = datetime.combine(event.event_date, time.min)
            appointment.End = datetime.combine(event.event_date + timedelta(days=1), time.min)
        else:
            start = _to_outlook_wall_datetime(event.event_date, event.start_time)
            end = _to_outlook_wall_datetime(event.event_date, event.end_time)
            if end < start:
                end += timedelta(days=1)
            if self._apply_event_time_zone(appointment, event.time_zone):
                self._apply_start_in_event_time_zone(appointment, start, end)
            else:
                appointment.Start = _outlook_datetime(start)
                appointment.End = _outlook_datetime(end)

        if event.url:
            try:
                appointment.WebPage = event.url
            except Exception:
                pass

        self._apply_invitee(appointment, invitee)

    def _apply_event_time_zone(self, appointment, iana_time_zone: str) -> bool:
        windows_time_zone = IANA_TO_WINDOWS_TIME_ZONE.get(iana_time_zone)
        try:
            outlook_time_zone = self.outlook.TimeZones.Item(windows_time_zone) if windows_time_zone else self.outlook.TimeZones.CurrentTimeZone
            appointment.StartTimeZone = outlook_time_zone
            appointment.EndTimeZone = outlook_time_zone
            return True
        except Exception:
            return False

    def _apply_start_in_event_time_zone(self, appointment, start: datetime, end: datetime) -> None:
        try:
            appointment.StartInStartTimeZone = _outlook_datetime(start)
            appointment.EndInEndTimeZone = _outlook_datetime(end)
        except Exception:
            appointment.Start = _outlook_datetime(start)
            appointment.End = _outlook_datetime(end)

    def _apply_invitee(self, appointment, invitee: str) -> None:
        clean_invitee = invitee.strip()
        if not clean_invitee:
            return
        try:
            appointment.MeetingStatus = OL_MEETING
            recipients = appointment.Recipients
            for index in range(1, recipients.Count + 1):
                recipient = recipients.Item(index)
                if recipient.Name.lower() == clean_invitee.lower() or getattr(recipient, "Address", "").lower() == clean_invitee.lower():
                    recipient.Type = OL_REQUIRED
                    recipients.ResolveAll()
                    return
            recipient = recipients.Add(clean_invitee)
            recipient.Type = OL_REQUIRED
            recipients.ResolveAll()
        except Exception:
            pass

    def _apply_reminder(self, appointment, minutes_before_start: int | None) -> None:
        if minutes_before_start is None:
            return
        try:
            appointment.ReminderSet = True
            appointment.ReminderMinutesBeforeStart = max(0, int(minutes_before_start))
        except Exception:
            pass


def _outlook_datetime(value: datetime) -> str:
    return value.strftime("%m/%d/%Y %I:%M %p")


def _to_outlook_wall_datetime(event_date: date, event_time: time) -> datetime:
    return datetime.combine(event_date, event_time)
