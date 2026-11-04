import sqlite3
from datetime import datetime, timedelta
from pathlib import Path

DB_PATH = Path(__file__).parent / "timesheet.db"


def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS timesheet_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            client_name TEXT NOT NULL,
            project_name TEXT NOT NULL,
            date TEXT NOT NULL,
            task_description TEXT NOT NULL,
            hours_worked REAL NOT NULL,
            start_time TEXT NOT NULL,
            end_time TEXT NOT NULL,
            remarks TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    conn.close()


def to_24h(time_str: str) -> str:
    """
    Convert time string to 24-hour HH:MM format.
    Accepts: '6:00 PM', '06:00 PM', '18:00', '6:00', '18:00:00' etc.
    """
    time_str = time_str.strip().upper().replace(".", ":")
    # Already 24h?
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            dt = datetime.strptime(time_str, fmt)
            return dt.strftime("%H:%M")
        except ValueError:
            pass
    # 12h formats
    for fmt in ("%I:%M %p", "%I:%M%p", "%I %p", "%I%p"):
        try:
            dt = datetime.strptime(time_str, fmt)
            return dt.strftime("%H:%M")
        except ValueError:
            pass
    raise ValueError(f"Invalid time format: {time_str}")


def to_12h(time_str: str) -> str:
    """Convert 24h HH:MM (or HH:MM:SS) to 12h like '6:00 PM'."""
    time_str = time_str.strip()
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            dt = datetime.strptime(time_str, fmt)
            return dt.strftime("%I:%M %p").lstrip("0")
        except ValueError:
            pass
    # Already 12h?
    return time_str


def calculate_hours(start_time_str: str, end_time_str: str) -> float:
    """Calculate hours between start and end time. Accepts 12h or 24h. Handles overnight."""
    try:
        start_24 = to_24h(start_time_str)
        end_24 = to_24h(end_time_str)
        start = datetime.strptime(start_24, "%H:%M")
        end = datetime.strptime(end_24, "%H:%M")

        # If end is before or equal start, assume next day
        if end <= start:
            end += timedelta(days=1)

        diff = end - start
        hours = diff.total_seconds() / 3600
        return round(hours, 2)
    except Exception:
        return 0.0


def add_entry(name, client_name, project_name, date, task_description,
              start_time, end_time, remarks=""):
    # Always store in 24h format
    start_24 = to_24h(start_time)
    end_24 = to_24h(end_time)
    hours = calculate_hours(start_24, end_24)
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO timesheet_entries
        (name, client_name, project_name, date, task_description,
         hours_worked, start_time, end_time, remarks)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (name, client_name, project_name, date, task_description,
          hours, start_24, end_24, remarks))
    conn.commit()
    entry_id = cursor.lastrowid
    conn.close()
    return entry_id, hours


def get_entries(start_date=None, end_date=None):
    conn = get_connection()
    cursor = conn.cursor()
    query = "SELECT * FROM timesheet_entries"
    params = []
    if start_date and end_date:
        query += " WHERE date BETWEEN ? AND ?"
        params = [start_date, end_date]
    query += " ORDER BY date, start_time"
    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()
    return [dict(row) for row in rows]


def get_entry_by_id(entry_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM timesheet_entries WHERE id = ?", (entry_id,))
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None


def get_entry_by_date(date_str: str):
    """Return the entry for a given date (YYYY-MM-DD), or None if none exists.
    Enforces one-entry-per-day rule.
    """
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT * FROM timesheet_entries WHERE date = ? ORDER BY id LIMIT 1",
        (date_str,)
    )
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None


def update_entry(entry_id, name, client_name, project_name, date,
                 task_description, start_time, end_time, remarks=""):
    start_24 = to_24h(start_time)
    end_24 = to_24h(end_time)
    hours = calculate_hours(start_24, end_24)
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE timesheet_entries
        SET name=?, client_name=?, project_name=?, date=?, task_description=?,
            hours_worked=?, start_time=?, end_time=?, remarks=?
        WHERE id=?
    """, (name, client_name, project_name, date, task_description,
          hours, start_24, end_24, remarks, entry_id))
    conn.commit()
    conn.close()
    return hours


def delete_entry(entry_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM timesheet_entries WHERE id = ?", (entry_id,))
    conn.commit()
    conn.close()


def get_week_dates(reference_date=None):
    """Return Monday to Friday dates of the week containing reference_date."""
    if reference_date is None:
        reference_date = datetime.now().date()
    elif isinstance(reference_date, str):
        reference_date = datetime.strptime(reference_date, "%Y-%m-%d").date()

    # Monday = 0
    monday = reference_date - timedelta(days=reference_date.weekday())
    week_days = [monday + timedelta(days=i) for i in range(5)]  # Mon-Fri
    return week_days


def get_total_hours(start_date=None, end_date=None):
    entries = get_entries(start_date, end_date)
    return round(sum(e["hours_worked"] for e in entries), 2)


def get_last_month_range(reference_date=None):
    """
    Return (start_date_str, end_date_str) for the full previous calendar month.
    e.g. if today is 2026-11-04 → ('2026-10-01', '2026-10-31')
    """
    if reference_date is None:
        reference_date = datetime.now().date()
    elif isinstance(reference_date, str):
        reference_date = datetime.strptime(reference_date, "%Y-%m-%d").date()

    # First day of current month
    first_of_this_month = reference_date.replace(day=1)
    # Last day of previous month
    last_of_prev = first_of_this_month - timedelta(days=1)
    # First day of previous month
    first_of_prev = last_of_prev.replace(day=1)

    return first_of_prev.strftime("%Y-%m-%d"), last_of_prev.strftime("%Y-%m-%d")