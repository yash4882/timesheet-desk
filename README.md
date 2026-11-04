# Consulting Timesheet Manager

Simple Windows desktop software for managing weekly consulting timesheets.

## New UI Layout (v2)

- **Form on TOP** – all fields visible, no scrolling needed
- **Entries table BELOW** – perfect for Mon–Fri (max 5 rows)
- Always-visible buttons: Calculate | Save Entry | Update | Clear | Delete
- Click any row → form fills → Update button activates

## Features

- Add / Edit / Delete timesheet entries
- Fields: Name | Client | Project | Date | Task | Hours | Start | End | Remarks
- Hours auto-calculate from Start → End time
- One entry per day rule
- **Week view** – Monday–Friday week + total hours
- **Last Month view** – full previous calendar month on one page + total hours
- Export weekly timesheet to Excel
- Export monthly timesheet (last completed month) to Excel

## How to Run

```bash
cd timesheet_app
pip install -r requirements.txt
python main.py
```

Or create .exe:
```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "TimesheetManager" main.py
```

## Usage

1. Fill form at the top (Name/Client/Project pre-filled)
2. Set Date, Task, Start Time, End Time → click **Calculate** or **Save Entry**
3. To edit: click a row in the table → form fills → click **Update**
4. Switch views:
   - **Week** → use Prev / This Week / Next to navigate weeks
   - **Last Month** → see all entries of the previous calendar month + total hours
5. Click **Download Weekly Excel** or **Download Monthly Excel** → files go to Downloads folder
