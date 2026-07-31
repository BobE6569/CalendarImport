# CalendarImport

CalendarImport is a standalone desktop program that imports itinerary rows from an Excel workbook into an Outlook or Google calendar.

## What It Does

- Asks you to select an Excel file.
- Reads the `Itinerary` tab.
- Uses the first row as column titles.
- Prompts for an event title prefix, defaulting to the Excel file name.
- Prompts for an optional user name or email to invite on each imported event.
- Creates or updates events in a selected Outlook or Google calendar.
- Uses the `Location` value as both the calendar event location and the suffix for the event title.
- Resolves event time zones from the location, reusing the previous row's time zone when a location is missing or cannot be resolved.
- Adds all remaining non-blank columns to the event notes as `Column title: value`, separated by one blank line.
- Adds the imported From / To date and time as the first line of the event notes.
- Always places the `Link` value at the bottom of the event notes.

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
- If an optional invitee is provided, CalendarImport adds that invitee to each event.

## Setup

1. Install dependencies:

```powershell
py -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
.\.venv\Scripts\python -m pip install -e .
```

2. Run the app:

```powershell
.\.venv\Scripts\python -m calendar_import
```

Or run `RunCalendarImport.cmd` / `RunCalendarImport.ps1`, which creates the virtual environment and installs dependencies automatically.

## Outlook Setup

Outlook import uses the installed Microsoft Outlook desktop app and the currently configured Outlook profile. Choose `Outlook`, connect, pick the target calendar folder, then import.

Outlook imports keep the clock time exactly as entered in Excel. For example, `7:30 AM` in the workbook is written as `7:30 AM` in Outlook.

The `Link` value is added to the appointment body and, when Outlook allows it, to the appointment web page field.

If Outlook shows a location as `Unknown`, that means Outlook could not match the typed location to one of its own place records. The location text is still saved on the event. Outlook's desktop automation interface does not provide a reliable way for CalendarImport to force-select one of the map suggestions.

## Google Setup

1. Create a Google Cloud OAuth desktop credential for the Google Calendar API.
2. Save the downloaded file as `credentials.json` in this project folder.
3. Choose `Google` in the app and connect.

The first Google run opens a Google sign-in page and stores `token.json` locally for future imports.

## Time Zones

CalendarImport includes a small starter map for common locations. If a location is unknown, the app asks for an IANA time zone such as `Europe/Rome` or `America/New_York`, then saves it in a workbook-specific text file such as `Norway_Itinerary_Aug_2026_Sub_timezones.txt`.

The text file uses this format:

```text
Location=IANA time zone
```

You can edit that file by hand if a location needs a different time zone later.

## Excel Permission Errors

If Windows reports `Permission denied` while reading a OneDrive workbook:

1. Close the workbook in Excel.
2. In File Explorer, right-click the workbook and choose **Always keep on this device**.
3. Try the import again.
4. If it still fails, save a copy to a local folder such as Documents and select that copy.
