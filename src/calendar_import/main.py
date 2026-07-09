from __future__ import annotations

from pathlib import Path
from tkinter import Button, Label, Listbox, StringVar, Tk, filedialog, messagebox, simpledialog

from .calendar_client import GoogleCalendarClient
from .excel_reader import ItineraryReadError, read_itinerary
from .timezone_resolver import DEFAULT_TIME_ZONE, TimeZoneResolver


PROJECT_DIR = Path(__file__).resolve().parents[2]


class CalendarImportApp:
    def __init__(self) -> None:
        self.root = Tk()
        self.root.title("CalendarImport")
        self.root.geometry("560x360")
        self.status = StringVar(value="Select an Excel itinerary file to begin.")
        self.excel_path: Path | None = None
        self.title_prefix = StringVar(value="")
        self.calendars: list[dict[str, str]] = []

        Label(self.root, text="CalendarImport", font=("Segoe UI", 18, "bold")).pack(pady=(18, 6))
        Label(self.root, textvariable=self.status, wraplength=500, justify="center").pack(pady=(0, 14))
        Button(self.root, text="Select Excel File", command=self.select_excel_file, width=24).pack(pady=4)
        Button(self.root, text="Connect Google Calendar", command=self.connect_calendar, width=24).pack(pady=4)
        Label(self.root, text="Target calendar").pack(pady=(14, 2))
        self.calendar_list = Listbox(self.root, height=6, exportselection=False)
        self.calendar_list.pack(fill="both", expand=True, padx=24, pady=4)
        Button(self.root, text="Import / Update Events", command=self.import_events, width=24).pack(pady=(8, 18))

    def run(self) -> None:
        self.root.mainloop()

    def select_excel_file(self) -> None:
        file_name = filedialog.askopenfilename(
            title="Select Excel itinerary",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if not file_name:
            return
        self.excel_path = Path(file_name)
        default_title = self.excel_path.stem
        title = simpledialog.askstring(
            "Event title",
            "Title to use for imported events:",
            initialvalue=default_title,
            parent=self.root,
        )
        self.title_prefix.set((title or default_title).strip() or default_title)
        self.status.set(f"Selected {self.excel_path.name}. Title: {self.title_prefix.get()}")

    def connect_calendar(self) -> None:
        try:
            client = GoogleCalendarClient(PROJECT_DIR)
            self.calendars = client.list_calendars()
        except Exception as exc:
            messagebox.showerror("Calendar connection failed", str(exc))
            return

        self.calendar_list.delete(0, "end")
        selected_index = 0
        for index, calendar in enumerate(self.calendars):
            label = calendar["summary"]
            if calendar["primary"]:
                label += " (primary)"
                selected_index = index
            self.calendar_list.insert("end", label)
        if self.calendars:
            self.calendar_list.selection_set(selected_index)
        self.status.set("Connected. Choose a target calendar, then import events.")

    def import_events(self) -> None:
        if not self.excel_path:
            messagebox.showwarning("Missing file", "Select an Excel file first.")
            return
        if not self.calendars:
            messagebox.showwarning("Missing calendar", "Connect to Google Calendar first.")
            return
        selection = self.calendar_list.curselection()
        if not selection:
            messagebox.showwarning("Missing calendar", "Choose the target calendar.")
            return

        resolver = TimeZoneResolver(PROJECT_DIR / "location_timezones.json", self.ask_for_time_zone)
        try:
            events = read_itinerary(self.excel_path, resolver)
        except ItineraryReadError as exc:
            messagebox.showerror("Excel import failed", str(exc))
            return
        if not events:
            messagebox.showinfo("No events", "No itinerary rows were found.")
            return

        client = GoogleCalendarClient(PROJECT_DIR)
        calendar_id = self.calendars[selection[0]]["id"]
        try:
            counts = client.upsert_events(calendar_id, self.title_prefix.get(), events, self.update_progress)
        except Exception as exc:
            messagebox.showerror("Calendar import failed", str(exc))
            return

        self.status.set(f"Done. Created {counts['created']} and updated {counts['updated']} events.")
        messagebox.showinfo("Import complete", self.status.get())

    def ask_for_time_zone(self, location: str, fallback: str = DEFAULT_TIME_ZONE) -> str | None:
        return simpledialog.askstring(
            "Time zone needed",
            f"Enter the IANA time zone for:\n{location}\n\nExample: Europe/Rome",
            initialvalue=fallback,
            parent=self.root,
        )

    def update_progress(self, index: int, total: int, action: str, summary: str) -> None:
        self.status.set(f"{index}/{total}: {action.capitalize()} {summary}")
        self.root.update_idletasks()


def main() -> None:
    CalendarImportApp().run()

