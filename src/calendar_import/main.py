from __future__ import annotations

from pathlib import Path
from tkinter import Button, Label, Listbox, Radiobutton, StringVar, Tk, filedialog, messagebox, simpledialog

from .calendar_client import GoogleCalendarClient
from .excel_reader import ItineraryReadError, read_itinerary
from .outlook_calendar_client import OutlookCalendarClient
from .preferences import Preferences
from .timezone_resolver import DEFAULT_TIME_ZONE, TimeZoneResolver, mapping_path_for_workbook


PROJECT_DIR = Path(__file__).resolve().parents[2]


class CalendarImportApp:
    def __init__(self) -> None:
        self.root = Tk()
        self.root.title("CalendarImport")
        self.root.geometry("560x430")
        self.status = StringVar(value="Select an Excel itinerary file to begin.")
        self.preferences = Preferences(PROJECT_DIR / "calendar_import_settings.json")
        self.excel_path: Path | None = None
        self.title_prefix = StringVar(value=self.preferences.get("last_title"))
        self.invitee = StringVar(value=self.preferences.get("last_invitee"))
        self.calendar_provider = StringVar(value="outlook")
        self.calendars: list[dict[str, str]] = []
        self.calendar_client = None

        Label(self.root, text="CalendarImport", font=("Segoe UI", 18, "bold")).pack(pady=(18, 6))
        Label(self.root, textvariable=self.status, wraplength=500, justify="center").pack(pady=(0, 14))
        Button(self.root, text="Select Excel File", command=self.select_excel_file, width=24).pack(pady=4)
        Label(self.root, text="Calendar service").pack(pady=(8, 0))
        Radiobutton(
            self.root,
            text="Outlook",
            variable=self.calendar_provider,
            value="outlook",
            command=self.clear_calendar_connection,
        ).pack()
        Radiobutton(
            self.root,
            text="Google",
            variable=self.calendar_provider,
            value="google",
            command=self.clear_calendar_connection,
        ).pack()
        Button(self.root, text="Connect Calendar", command=self.connect_calendar, width=24).pack(pady=4)
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
        default_title = self.title_prefix.get().strip() or self.excel_path.stem
        title = simpledialog.askstring(
            "Event title",
            "Title to use for imported events:",
            initialvalue=default_title,
            parent=self.root,
        )
        self.title_prefix.set((title or default_title).strip() or default_title)
        self.preferences.set("last_title", self.title_prefix.get())
        invitee = simpledialog.askstring(
            "Optional invitee",
            "User name or email to invite on each calendar event:",
            initialvalue=self.invitee.get(),
            parent=self.root,
        )
        self.invitee.set((invitee or "").strip())
        self.preferences.set("last_invitee", self.invitee.get())
        invitee_status = f" Invitee: {self.invitee.get()}" if self.invitee.get() else ""
        self.status.set(f"Selected {self.excel_path.name}. Title: {self.title_prefix.get()}.{invitee_status}")

    def clear_calendar_connection(self) -> None:
        self.calendar_client = None
        self.calendars = []
        self.calendar_list.delete(0, "end")
        self.status.set(f"Selected {self.calendar_provider.get().capitalize()}. Connect to choose a calendar.")

    def connect_calendar(self) -> None:
        try:
            self.calendar_client = self._create_calendar_client()
            self.calendars = self.calendar_client.list_calendars()
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
        self.status.set(f"Connected to {self.calendar_provider.get().capitalize()}. Choose a target calendar, then import events.")

    def import_events(self) -> None:
        if not self.excel_path:
            messagebox.showwarning("Missing file", "Select an Excel file first.")
            return
        if not self.calendars or not self.calendar_client:
            messagebox.showwarning("Missing calendar", "Connect to a calendar first.")
            return
        selection = self.calendar_list.curselection()
        if not selection:
            messagebox.showwarning("Missing calendar", "Choose the target calendar.")
            return

        resolver = TimeZoneResolver(mapping_path_for_workbook(PROJECT_DIR, self.excel_path), self.ask_for_time_zone)
        try:
            events = read_itinerary(self.excel_path, resolver)
        except ItineraryReadError as exc:
            messagebox.showerror("Excel import failed", str(exc))
            return
        if not events:
            messagebox.showinfo("No events", "No itinerary rows were found.")
            return

        calendar_id = self.calendars[selection[0]]["id"]
        try:
            counts = self.calendar_client.upsert_events(
                calendar_id,
                self.title_prefix.get(),
                events,
                self.invitee.get(),
                self.update_progress,
            )
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

    def _create_calendar_client(self):
        if self.calendar_provider.get() == "outlook":
            return OutlookCalendarClient()
        return GoogleCalendarClient(PROJECT_DIR)


def main() -> None:
    CalendarImportApp().run()
