# CalendarImport

CalendarImport is a standalone desktop program that imports itinerary rows from an Excel workbook into a Google Calendar.

## What It Does

- Asks you to select an Excel file.
- Reads the `Itinerary` tab.
- Uses the first row as column titles.
- Prompts for an event title prefix, defaulting to the Excel file name.
- Creates or updates events in a selected Google Calendar.
- Uses the `Location` value as both the calendar event location and the suffix for the event title.
- Resolves event time zones from the location, reusing the previous row's time zone when a location is missing or cannot be resolved.
- Adds all remaining non-blank columns to the event notes as `Column title: value`.

## Excel Layout

The `Itinerary` sheet must start with these first five columns:

| Column | Meaning |
| --- | --- |
| Date | Calendar event date |
| Location | Event location, title suffix, and time-zone lookup value |
| From | Start time, or blank |
| To | End time, or blank |
| Link | Optional event URL |

Any columns after `Link` are added to the event notes when their row value is non-blank.

## Calendar Event Rules

- Event title is `<title prefix> - <location>` when location exists, otherwise just `<title prefix>`.
- If `From` and `To` are both blank, the event is an all-day event.
- If only one of `From` or `To` is filled, start and end are set to that same time.
- If both are filled, start uses `From` and end uses `To`.
- Existing events are found by exact title in the selected calendar. When several match, CalendarImport prefers the one on the same date.

## Setup

1. Create a Google Cloud OAuth desktop credential for the Google Calendar API.
2. Save the downloaded file as `credentials.json` in this project folder.
3. Install dependencies:

```powershell
py -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
```

4. Run the app:

```powershell
.\.venv\Scripts\python -m calendar_import
```

Or run `RunCalendarImport.ps1`, which creates the virtual environment and installs dependencies automatically.

The first run opens a Google sign-in page and stores `token.json` locally for future imports.

## Time Zones

CalendarImport includes a small starter map for common locations. If a location is unknown, the app asks for an IANA time zone such as `Europe/Rome` or `America/New_York`, then saves it in `location_timezones.json`.
