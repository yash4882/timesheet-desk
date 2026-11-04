import customtkinter as ctk
from tkinter import messagebox, ttk
from datetime import datetime, timedelta
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

from database import (
    init_db, add_entry, get_entries, update_entry, delete_entry,
    calculate_hours, get_week_dates, get_total_hours, get_entry_by_id,
    get_entry_by_date, to_12h, to_24h, get_last_month_range
)

try:
    from tkcalendar import DateEntry
    HAS_TKCALENDAR = True
except ImportError:
    HAS_TKCALENDAR = False

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

DEFAULT_NAME = "Harshita Ugde"
DEFAULT_CLIENT = "Zeus"
DEFAULT_PROJECT = "Ignition"


class TimesheetApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Consulting Timesheet Manager")
        self.geometry("1100x820")
        self.minsize(980, 720)

        init_db()

        self.selected_entry_id = None
        self.view_mode = "week"  # "week" or "month"
        self.create_widgets()
        self.refresh_table()

    # ------------------------------------------------------------------
    # Time helpers
    # ------------------------------------------------------------------
    def _build_time_row(self, parent, default_hour="6", default_minute="00", default_ampm="PM"):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        hours = [str(i) for i in range(1, 13)]
        minutes = [f"{i:02d}" for i in range(0, 60, 5)]
        ampm = ["AM", "PM"]

        hour_cb = ctk.CTkComboBox(frame, values=hours, width=64, state="readonly")
        hour_cb.set(default_hour)
        hour_cb.pack(side="left", padx=(0, 3))

        ctk.CTkLabel(frame, text=":", font=ctk.CTkFont(size=15, weight="bold")).pack(side="left")

        min_cb = ctk.CTkComboBox(frame, values=minutes, width=64, state="readonly")
        min_cb.set(default_minute)
        min_cb.pack(side="left", padx=(3, 6))

        ampm_cb = ctk.CTkComboBox(frame, values=ampm, width=64, state="readonly")
        ampm_cb.set(default_ampm)
        ampm_cb.pack(side="left")

        return frame, hour_cb, min_cb, ampm_cb

    def _get_time_str(self, hour_cb, min_cb, ampm_cb) -> str:
        return f"{hour_cb.get().strip()}:{min_cb.get().strip()} {ampm_cb.get().strip()}"

    def _set_time_widgets(self, hour_cb, min_cb, ampm_cb, time_24h: str):
        try:
            dt = datetime.strptime(time_24h[:5], "%H:%M")
            hour_12 = dt.strftime("%I").lstrip("0") or "12"
            minute = f"{(int(dt.strftime('%M')) // 5) * 5:02d}"
            ampm = dt.strftime("%p")
            hour_cb.set(hour_12)
            min_cb.set(minute)
            ampm_cb.set(ampm)
        except Exception:
            hour_cb.set("6")
            min_cb.set("00")
            ampm_cb.set("PM")

    # ------------------------------------------------------------------
    def create_widgets(self):
        # ==================== TOP: ENTRY FORM ====================
        form_card = ctk.CTkFrame(self, corner_radius=10)
        form_card.pack(fill="x", padx=12, pady=(12, 6))

        # Title row
        title_row = ctk.CTkFrame(form_card, fg_color="transparent")
        title_row.pack(fill="x", padx=14, pady=(10, 4))
        ctk.CTkLabel(
            title_row, text="Add / Edit Entry",
            font=ctk.CTkFont(size=17, weight="bold")
        ).pack(side="left")
        self.hours_label = ctk.CTkLabel(
            title_row, text="Hours: 0.00",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#27ae60"
        )
        self.hours_label.pack(side="right")

        # ---- Row 1: Name | Client | Project | Date ----
        row1 = ctk.CTkFrame(form_card, fg_color="transparent")
        row1.pack(fill="x", padx=14, pady=(4, 2))

        # Name
        f = ctk.CTkFrame(row1, fg_color="transparent")
        f.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkLabel(f, text="Name", anchor="w", font=ctk.CTkFont(size=12)).pack(fill="x")
        self.name_entry = ctk.CTkEntry(f, height=32)
        self.name_entry.insert(0, DEFAULT_NAME)
        self.name_entry.pack(fill="x", pady=(2, 0))

        # Client
        f = ctk.CTkFrame(row1, fg_color="transparent")
        f.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkLabel(f, text="Client Name", anchor="w", font=ctk.CTkFont(size=12)).pack(fill="x")
        self.client_entry = ctk.CTkEntry(f, height=32)
        self.client_entry.insert(0, DEFAULT_CLIENT)
        self.client_entry.pack(fill="x", pady=(2, 0))

        # Project
        f = ctk.CTkFrame(row1, fg_color="transparent")
        f.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkLabel(f, text="Project Name", anchor="w", font=ctk.CTkFont(size=12)).pack(fill="x")
        self.project_entry = ctk.CTkEntry(f, height=32)
        self.project_entry.insert(0, DEFAULT_PROJECT)
        self.project_entry.pack(fill="x", pady=(2, 0))

        # Date
        f = ctk.CTkFrame(row1, fg_color="transparent")
        f.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(f, text="Date", anchor="w", font=ctk.CTkFont(size=12)).pack(fill="x")
        if HAS_TKCALENDAR:
            self.date_picker = DateEntry(
                f, width=14, background="#1f6aa5", foreground="white",
                borderwidth=2, date_pattern="yyyy-mm-dd", font=("Segoe UI", 11)
            )
            self.date_picker.set_date(datetime.now().date())
            self.date_picker.pack(fill="x", pady=(2, 0))
        else:
            self.date_picker = ctk.CTkEntry(f, height=32)
            self.date_picker.insert(0, datetime.now().strftime("%Y-%m-%d"))
            self.date_picker.pack(fill="x", pady=(2, 0))

        # ---- Row 2: Task Description (full width) ----
        row2 = ctk.CTkFrame(form_card, fg_color="transparent")
        row2.pack(fill="x", padx=14, pady=(6, 2))
        ctk.CTkLabel(row2, text="Task Description", anchor="w", font=ctk.CTkFont(size=12)).pack(fill="x")
        self.task_entry = ctk.CTkTextbox(row2, height=56, wrap="word")
        self.task_entry.pack(fill="x", pady=(2, 0))

        # ---- Row 3: Start Time | End Time | Remarks | Buttons ----
        row3 = ctk.CTkFrame(form_card, fg_color="transparent")
        row3.pack(fill="x", padx=14, pady=(6, 10))

        # Start Time
        f = ctk.CTkFrame(row3, fg_color="transparent")
        f.pack(side="left", padx=(0, 12))
        ctk.CTkLabel(f, text="Start Time", anchor="w", font=ctk.CTkFont(size=12)).pack(anchor="w")
        time_frame, self.start_hour, self.start_min, self.start_ampm = self._build_time_row(
            f, default_hour="6", default_minute="00", default_ampm="PM"
        )
        time_frame.pack(anchor="w", pady=(2, 0))

        # End Time
        f = ctk.CTkFrame(row3, fg_color="transparent")
        f.pack(side="left", padx=(0, 12))
        ctk.CTkLabel(f, text="End Time", anchor="w", font=ctk.CTkFont(size=12)).pack(anchor="w")
        time_frame, self.end_hour, self.end_min, self.end_ampm = self._build_time_row(
            f, default_hour="10", default_minute="00", default_ampm="PM"
        )
        time_frame.pack(anchor="w", pady=(2, 0))

        # Remarks
        f = ctk.CTkFrame(row3, fg_color="transparent")
        f.pack(side="left", fill="x", expand=True, padx=(0, 12))
        ctk.CTkLabel(f, text="Remarks (if any)", anchor="w", font=ctk.CTkFont(size=12)).pack(fill="x")
        self.remarks_entry = ctk.CTkEntry(f, height=32)
        self.remarks_entry.pack(fill="x", pady=(2, 0))

        # Buttons
        btn_box = ctk.CTkFrame(row3, fg_color="transparent")
        btn_box.pack(side="right")

        ctk.CTkButton(
            btn_box, text="Calculate", width=90, height=30,
            command=self.calculate_hours_ui
        ).pack(side="left", padx=2)

        ctk.CTkButton(
            btn_box, text="Save Entry", width=100, height=30,
            fg_color="#27ae60", hover_color="#1e8449",
            command=self.save_entry
        ).pack(side="left", padx=2)

        ctk.CTkButton(
            btn_box, text="Update", width=90, height=30,
            fg_color="#2980b9", hover_color="#1c6ea4",
            command=self.update_entry_ui
        ).pack(side="left", padx=2)

        ctk.CTkButton(
            btn_box, text="Clear", width=70, height=30,
            command=self.clear_form
        ).pack(side="left", padx=2)

        ctk.CTkButton(
            btn_box, text="Delete", width=70, height=30,
            fg_color="#c0392b", hover_color="#a93226",
            command=self.delete_entry_ui
        ).pack(side="left", padx=2)

        # ==================== BOTTOM: TABLE (Week / Month) ====================
        table_card = ctk.CTkFrame(self, corner_radius=10)
        table_card.pack(fill="both", expand=True, padx=12, pady=(6, 12))

        # ---- Bottom bar FIRST (side=bottom) so it is ALWAYS visible ----
        bottom_bar = ctk.CTkFrame(table_card, fg_color="transparent")
        bottom_bar.pack(side="bottom", fill="x", padx=12, pady=(6, 10))

        self.total_label = ctk.CTkLabel(
            bottom_bar, text="Total Hours (week): 0.00",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#27ae60"
        )
        self.total_label.pack(side="left")

        ctk.CTkButton(
            bottom_bar, text="Download Monthly Excel", width=170, height=34,
            fg_color="#8e44ad", hover_color="#7d3c98",
            command=self.export_month_excel
        ).pack(side="right", padx=4)

        ctk.CTkButton(
            bottom_bar, text="Download Weekly Excel", width=160, height=34,
            fg_color="#2980b9", command=self.export_week_excel
        ).pack(side="right", padx=4)

        # ---- View mode + navigation controls (top of table card) ----
        week_frame = ctk.CTkFrame(table_card, fg_color="transparent")
        week_frame.pack(side="top", fill="x", padx=12, pady=(10, 4))

        ctk.CTkLabel(week_frame, text="View:", font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")

        self.btn_week_view = ctk.CTkButton(
            week_frame, text="Week", width=70, height=28,
            fg_color="#1f6aa5", hover_color="#144870",
            command=self.set_week_view
        )
        self.btn_week_view.pack(side="left", padx=(4, 2))

        self.btn_month_view = ctk.CTkButton(
            week_frame, text="Last Month", width=100, height=28,
            fg_color="#555555", hover_color="#444444",
            command=self.set_month_view
        )
        self.btn_month_view.pack(side="left", padx=2)

        # Week navigation (only useful in week view)
        self.week_nav_frame = ctk.CTkFrame(week_frame, fg_color="transparent")
        self.week_nav_frame.pack(side="left", padx=(12, 0))

        ctk.CTkButton(self.week_nav_frame, text="◀ Prev", width=70, height=28, command=self.prev_week).pack(side="left", padx=2)
        ctk.CTkButton(self.week_nav_frame, text="This Week", width=80, height=28, command=self.load_this_week).pack(side="left", padx=2)
        ctk.CTkButton(self.week_nav_frame, text="Next ▶", width=70, height=28, command=self.next_week).pack(side="left", padx=2)

        self.week_date_var = ctk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.week_entry = ctk.CTkEntry(self.week_nav_frame, textvariable=self.week_date_var, width=100, height=28)
        self.week_entry.pack(side="left", padx=4)
        ctk.CTkButton(self.week_nav_frame, text="Go", width=45, height=28, command=self.load_week).pack(side="left", padx=2)

        self.week_range_label = ctk.CTkLabel(
            week_frame, text="", font=ctk.CTkFont(size=12), text_color="#7f8c8d"
        )
        self.week_range_label.pack(side="left", padx=10)

        # ---- Table fills remaining space between top controls and bottom bar ----
        table_frame = ctk.CTkFrame(table_card, fg_color="transparent")
        table_frame.pack(side="top", fill="both", expand=True, padx=10, pady=(0, 4))

        columns = ("id", "date", "client", "project", "hours", "start", "end", "task")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=10)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", rowheight=30, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        style.map("Treeview", background=[("selected", "#1f6aa5")])

        self.tree.heading("id", text="ID")
        self.tree.heading("date", text="Date")
        self.tree.heading("client", text="Client")
        self.tree.heading("project", text="Project")
        self.tree.heading("hours", text="Hours")
        self.tree.heading("start", text="Start")
        self.tree.heading("end", text="End")
        self.tree.heading("task", text="Task Description")

        self.tree.column("id", width=40, anchor="center", stretch=False)
        self.tree.column("date", width=95, anchor="center", stretch=False)
        self.tree.column("client", width=80, anchor="center", stretch=False)
        self.tree.column("project", width=80, anchor="center", stretch=False)
        self.tree.column("hours", width=60, anchor="center", stretch=False)
        self.tree.column("start", width=85, anchor="center", stretch=False)
        self.tree.column("end", width=85, anchor="center", stretch=False)
        self.tree.column("task", width=320)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.tree.bind("<<TreeviewSelect>>", self.on_select)

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def get_current_week_range(self):
        try:
            ref = datetime.strptime(self.week_date_var.get().strip(), "%Y-%m-%d").date()
        except ValueError:
            ref = datetime.now().date()
        days = get_week_dates(ref)
        return days[0].strftime("%Y-%m-%d"), days[-1].strftime("%Y-%m-%d")

    def calculate_hours_ui(self):
        start = self._get_time_str(self.start_hour, self.start_min, self.start_ampm)
        end = self._get_time_str(self.end_hour, self.end_min, self.end_ampm)
        hours = calculate_hours(start, end)
        self.hours_label.configure(text=f"Hours: {hours:.2f}")

    def clear_form(self):
        self.selected_entry_id = None
        self.name_entry.delete(0, "end")
        self.name_entry.insert(0, DEFAULT_NAME)
        self.client_entry.delete(0, "end")
        self.client_entry.insert(0, DEFAULT_CLIENT)
        self.project_entry.delete(0, "end")
        self.project_entry.insert(0, DEFAULT_PROJECT)
        if HAS_TKCALENDAR:
            self.date_picker.set_date(datetime.now().date())
        else:
            self.date_picker.delete(0, "end")
            self.date_picker.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.task_entry.delete("1.0", "end")
        self.remarks_entry.delete(0, "end")
        self.start_hour.set("6")
        self.start_min.set("00")
        self.start_ampm.set("PM")
        self.end_hour.set("10")
        self.end_min.set("00")
        self.end_ampm.set("PM")
        self.hours_label.configure(text="Hours: 0.00")

    def get_form_data(self):
        name = self.name_entry.get().strip()
        client = self.client_entry.get().strip()
        project = self.project_entry.get().strip()
        if HAS_TKCALENDAR:
            date_str = self.date_picker.get_date().strftime("%Y-%m-%d")
        else:
            date_str = self.date_picker.get().strip()
        task = self.task_entry.get("1.0", "end").strip()
        start = self._get_time_str(self.start_hour, self.start_min, self.start_ampm)
        end = self._get_time_str(self.end_hour, self.end_min, self.end_ampm)
        remarks = self.remarks_entry.get().strip()
        return name, client, project, date_str, task, start, end, remarks

    def save_entry(self):
        name, client, project, date_str, task, start, end, remarks = self.get_form_data()
        if not all([name, client, project, date_str, task]):
            messagebox.showwarning("Missing Data", "Please fill Name, Client, Project, Date and Task Description.")
            return
        existing = get_entry_by_date(date_str)
        if existing:
            messagebox.showwarning("Duplicate", f"An entry already exists for {date_str}. Use Update instead.")
            return
        try:
            entry_id, hours = add_entry(name, client, project, date_str, task, start, end, remarks)
            self.hours_label.configure(text=f"Hours: {hours:.2f}")
            messagebox.showinfo("Saved", f"Entry saved successfully (ID: {entry_id})")
            self.clear_form()
            self.refresh_table()
        except Exception as ex:
            messagebox.showerror("Error", str(ex))

    def update_entry_ui(self):
        if not self.selected_entry_id:
            messagebox.showwarning("No Selection", "Please select an entry from the table to update.")
            return
        name, client, project, date_str, task, start, end, remarks = self.get_form_data()
        if not all([name, client, project, date_str, task]):
            messagebox.showwarning("Missing Data", "Please fill all required fields.")
            return
        try:
            update_entry(self.selected_entry_id, name, client, project, date_str, task, start, end, remarks)
            messagebox.showinfo("Updated", "Entry updated successfully.")
            self.clear_form()
            self.refresh_table()
        except Exception as ex:
            messagebox.showerror("Error", str(ex))

    def delete_entry_ui(self):
        if not self.selected_entry_id:
            messagebox.showwarning("No Selection", "Please select an entry to delete.")
            return
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this entry?"):
            delete_entry(self.selected_entry_id)
            self.clear_form()
            self.refresh_table()

    def on_select(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        item = self.tree.item(selected[0])
        entry_id = item["values"][0]
        entry = get_entry_by_id(entry_id)
        if not entry:
            return
        self.selected_entry_id = entry_id
        self.name_entry.delete(0, "end")
        self.name_entry.insert(0, entry["name"])
        self.client_entry.delete(0, "end")
        self.client_entry.insert(0, entry["client_name"])
        self.project_entry.delete(0, "end")
        self.project_entry.insert(0, entry["project_name"])
        if HAS_TKCALENDAR:
            self.date_picker.set_date(datetime.strptime(entry["date"], "%Y-%m-%d").date())
        else:
            self.date_picker.delete(0, "end")
            self.date_picker.insert(0, entry["date"])
        self.task_entry.delete("1.0", "end")
        self.task_entry.insert("1.0", entry["task_description"])
        self.remarks_entry.delete(0, "end")
        self.remarks_entry.insert(0, entry["remarks"] or "")
        self._set_time_widgets(self.start_hour, self.start_min, self.start_ampm, entry["start_time"])
        self._set_time_widgets(self.end_hour, self.end_min, self.end_ampm, entry["end_time"])
        self.hours_label.configure(text=f"Hours: {entry['hours_worked']:.2f}")

    def set_week_view(self):
        self.view_mode = "week"
        self.btn_week_view.configure(fg_color="#1f6aa5", hover_color="#144870")
        self.btn_month_view.configure(fg_color="#555555", hover_color="#444444")
        # Show week navigation
        self.week_nav_frame.pack(side="left", padx=(12, 0))
        self.refresh_table()

    def set_month_view(self):
        self.view_mode = "month"
        self.btn_month_view.configure(fg_color="#8e44ad", hover_color="#7d3c98")
        self.btn_week_view.configure(fg_color="#555555", hover_color="#444444")
        # Hide week navigation (not needed for month view)
        self.week_nav_frame.pack_forget()
        self.refresh_table()

    def load_this_week(self):
        self.view_mode = "week"
        self.btn_week_view.configure(fg_color="#1f6aa5", hover_color="#144870")
        self.btn_month_view.configure(fg_color="#555555", hover_color="#444444")
        self.week_nav_frame.pack(side="left", padx=(12, 0))
        self.week_date_var.set(datetime.now().strftime("%Y-%m-%d"))
        self.refresh_table()

    def prev_week(self):
        try:
            d = datetime.strptime(self.week_date_var.get(), "%Y-%m-%d").date()
            d = d - timedelta(days=7)
            self.week_date_var.set(d.strftime("%Y-%m-%d"))
            self.view_mode = "week"
            self.refresh_table()
        except Exception:
            pass

    def next_week(self):
        try:
            d = datetime.strptime(self.week_date_var.get(), "%Y-%m-%d").date()
            d = d + timedelta(days=7)
            self.week_date_var.set(d.strftime("%Y-%m-%d"))
            self.view_mode = "week"
            self.refresh_table()
        except Exception:
            pass

    def load_week(self):
        self.view_mode = "week"
        self.refresh_table()

    def refresh_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        if self.view_mode == "month":
            start, end = get_last_month_range()
            entries = get_entries(start, end)

            try:
                s = datetime.strptime(start, "%Y-%m-%d").date()
                e = datetime.strptime(end, "%Y-%m-%d").date()
                month_name = s.strftime("%B %Y")
                self.week_range_label.configure(
                    text=f"Showing: {s.strftime('%d %b %Y')}  →  {e.strftime('%d %b %Y')}  ({month_name})"
                )
            except Exception:
                self.week_range_label.configure(text="")

            for e in entries:
                self.tree.insert("", "end", values=(
                    e["id"],
                    e["date"],
                    e["client_name"],
                    e["project_name"],
                    f"{e['hours_worked']:.2f}",
                    to_12h(e["start_time"]),
                    to_12h(e["end_time"]),
                    e["task_description"][:90] + ("..." if len(e["task_description"]) > 90 else "")
                ))

            total = get_total_hours(start, end)
            self.total_label.configure(
                text=f"Total Hours (last month): {total:.2f}",
                text_color="#8e44ad"
            )
        else:
            # Week view
            start, end = self.get_current_week_range()
            entries = get_entries(start, end)

            try:
                s = datetime.strptime(start, "%Y-%m-%d").date()
                e = datetime.strptime(end, "%Y-%m-%d").date()
                self.week_range_label.configure(
                    text=f"Showing: {s.strftime('%d %b %Y')}  →  {e.strftime('%d %b %Y')}  (Mon–Fri)"
                )
            except Exception:
                self.week_range_label.configure(text="")

            for e in entries:
                self.tree.insert("", "end", values=(
                    e["id"],
                    e["date"],
                    e["client_name"],
                    e["project_name"],
                    f"{e['hours_worked']:.2f}",
                    to_12h(e["start_time"]),
                    to_12h(e["end_time"]),
                    e["task_description"][:90] + ("..." if len(e["task_description"]) > 90 else "")
                ))

            total = get_total_hours(start, end)
            self.total_label.configure(
                text=f"Total Hours (week): {total:.2f}",
                text_color="#27ae60"
            )

    def export_week_excel(self):
        start, end = self.get_current_week_range()
        entries = get_entries(start, end)

        if not entries:
            messagebox.showwarning("No Data", "No entries found for this week.")
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Timesheet"

        header_font = Font(name="Aptos Narrow", size=11, bold=True, color="000000")
        header_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
        data_font = Font(name="Aptos Narrow", size=11)
        thin_border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin")
        )
        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

        headers = [
            "Name", "Client Name", "Project Name", "Date", "Task Description",
            "Hours Worked", "Start Time", "End Time", "Remarks (if any)"
        ]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align
            cell.border = thin_border

        for row_idx, e in enumerate(entries, 2):
            ws.cell(row=row_idx, column=1, value=e["name"]).font = data_font
            ws.cell(row=row_idx, column=2, value=e["client_name"]).font = data_font
            ws.cell(row=row_idx, column=3, value=e["project_name"]).font = data_font

            try:
                date_obj = datetime.strptime(e["date"], "%Y-%m-%d")
            except Exception:
                date_obj = e["date"]
            cell = ws.cell(row=row_idx, column=4, value=date_obj)
            cell.font = data_font
            cell.number_format = "YYYY-MM-DD"

            ws.cell(row=row_idx, column=5, value=e["task_description"]).font = data_font

            hours_val = e["hours_worked"]
            hours_str = f"{int(hours_val)} hours" if hours_val == int(hours_val) else f"{hours_val} hours"
            ws.cell(row=row_idx, column=6, value=hours_str).font = data_font

            # 12-hour format (e.g. "6:00 PM")
            start_12 = to_12h(e["start_time"])
            end_12 = to_12h(e["end_time"])
            ws.cell(row=row_idx, column=7, value=start_12).font = data_font
            ws.cell(row=row_idx, column=8, value=end_12).font = data_font

            ws.cell(row=row_idx, column=9, value=e["remarks"] or "").font = data_font

            for col in range(1, 10):
                cell = ws.cell(row=row_idx, column=col)
                cell.border = thin_border
                cell.alignment = center_align if col in (2, 3, 4, 6, 7, 8) else left_align

        # Fixed column widths
        widths = [15, 13, 14, 12, 55, 13, 12, 12, 22]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.row_dimensions[1].height = 22
        for r in range(2, len(entries) + 2):
            ws.row_dimensions[r].height = 45

        downloads = Path.home() / "Downloads"
        downloads.mkdir(exist_ok=True)
        filename = f"Consulting_Timesheet_{start}_to_{end}.xlsx"
        filepath = downloads / filename
        wb.save(filepath)

        exports_dir = Path(__file__).parent / "exports"
        exports_dir.mkdir(exist_ok=True)
        wb.save(exports_dir / filename)

        messagebox.showinfo("Exported", f"Weekly timesheet saved to:\n{filepath}")

    def export_month_excel(self):
        """Export timesheet for the completed previous calendar month."""
        start, end = get_last_month_range()
        entries = get_entries(start, end)

        if not entries:
            # Show friendly month name in the message
            try:
                month_name = datetime.strptime(start, "%Y-%m-%d").strftime("%B %Y")
            except Exception:
                month_name = f"{start} to {end}"
            messagebox.showwarning(
                "No Data",
                f"No entries found for the last completed month ({month_name})."
            )
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Timesheet"

        header_font = Font(name="Aptos Narrow", size=11, bold=True, color="000000")
        header_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
        data_font = Font(name="Aptos Narrow", size=11)
        thin_border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin")
        )
        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

        headers = [
            "Name", "Client Name", "Project Name", "Date", "Task Description",
            "Hours Worked", "Start Time", "End Time", "Remarks (if any)"
        ]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align
            cell.border = thin_border

        for row_idx, e in enumerate(entries, 2):
            ws.cell(row=row_idx, column=1, value=e["name"]).font = data_font
            ws.cell(row=row_idx, column=2, value=e["client_name"]).font = data_font
            ws.cell(row=row_idx, column=3, value=e["project_name"]).font = data_font

            try:
                date_obj = datetime.strptime(e["date"], "%Y-%m-%d")
            except Exception:
                date_obj = e["date"]
            cell = ws.cell(row=row_idx, column=4, value=date_obj)
            cell.font = data_font
            cell.number_format = "YYYY-MM-DD"

            ws.cell(row=row_idx, column=5, value=e["task_description"]).font = data_font

            hours_val = e["hours_worked"]
            hours_str = f"{int(hours_val)} hours" if hours_val == int(hours_val) else f"{hours_val} hours"
            ws.cell(row=row_idx, column=6, value=hours_str).font = data_font

            # 12-hour format (e.g. "6:00 PM")
            start_12 = to_12h(e["start_time"])
            end_12 = to_12h(e["end_time"])
            ws.cell(row=row_idx, column=7, value=start_12).font = data_font
            ws.cell(row=row_idx, column=8, value=end_12).font = data_font

            ws.cell(row=row_idx, column=9, value=e["remarks"] or "").font = data_font

            for col in range(1, 10):
                cell = ws.cell(row=row_idx, column=col)
                cell.border = thin_border
                cell.alignment = center_align if col in (2, 3, 4, 6, 7, 8) else left_align

        # Fixed column widths
        widths = [15, 13, 14, 12, 55, 13, 12, 12, 22]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.row_dimensions[1].height = 22
        for r in range(2, len(entries) + 2):
            ws.row_dimensions[r].height = 45

        downloads = Path.home() / "Downloads"
        downloads.mkdir(exist_ok=True)
        filename = f"Consulting_Timesheet_Monthly_{start}_to_{end}.xlsx"
        filepath = downloads / filename
        wb.save(filepath)

        exports_dir = Path(__file__).parent / "exports"
        exports_dir.mkdir(exist_ok=True)
        wb.save(exports_dir / filename)

        try:
            month_name = datetime.strptime(start, "%Y-%m-%d").strftime("%B %Y")
        except Exception:
            month_name = f"{start} to {end}"

        messagebox.showinfo(
            "Exported",
            f"Monthly timesheet for {month_name} saved to:\n{filepath}"
        )


if __name__ == "__main__":
    app = TimesheetApp()
    app.mainloop()