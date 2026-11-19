"""
Excel download log for the SkillPrism Reports Downloader.

Every report attempt (success or failure) is appended as one row to
    <download root>\\Download_Log.xlsx
with three sheets:
    Log      - one row per report attempt (filterable table)
    Summary  - per client + report: attempts, successes, success rate, last success, avg time
    By Day   - per day: attempts, successes, failures, total MB downloaded

If the workbook is open in Excel when a run finishes (Windows locks it), the
row is parked in pending_log.csv next to it and merged in on the next write,
so nothing is lost.
"""
import csv
from collections import defaultdict
from datetime import datetime
from pathlib import Path

COLUMNS = [
    "Run ID", "Date", "Day", "Trigger", "Client", "Admin User", "Report", "Status",
    "Started At", "Finished At", "Duration (min)", "File Name", "File Size (MB)",
    "Folder", "Error",
]
WIDTHS = [17, 12, 6, 11, 22, 22, 28, 10, 10, 10, 9, 34, 10, 45, 60]


def _pending_path(log_path: Path) -> Path:
    return log_path.with_name("pending_log.csv")


def make_entry(run_id, trigger, client, admin_user, report, status, started, finished,
               file_path=None, error=None):
    size_mb = None
    if file_path and Path(file_path).exists():
        size_mb = round(Path(file_path).stat().st_size / (1024 * 1024), 2)
    return {
        "Run ID": run_id,
        "Date": started.strftime("%d-%m-%Y"),
        "Day": started.strftime("%a"),
        "Trigger": trigger,
        "Client": client,
        "Admin User": admin_user,
        "Report": report,
        "Status": status,
        "Started At": started.strftime("%H:%M:%S"),
        "Finished At": finished.strftime("%H:%M:%S"),
        "Duration (min)": round((finished - started).total_seconds() / 60, 1),
        "File Name": Path(file_path).name if file_path else "",
        "File Size (MB)": size_mb,
        "Folder": str(Path(file_path).parent) if file_path else "",
        "Error": (error or "")[:500],
    }


def append_entry(log_path: Path, entry: dict, log=lambda m: None):
    log_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        _append_csv(_pending_path(log_path), [entry])
        log("Excel log: openpyxl isn't installed (run 1-Setup.bat again) - wrote to "
            "pending_log.csv instead.")
        return

    pending = _read_pending(log_path)
    try:
        _write_workbook(log_path, pending + [entry])
        if pending:
            _pending_path(log_path).unlink(missing_ok=True)
        log(f"Excel log updated: {log_path}")
    except PermissionError:
        _append_csv(_pending_path(log_path), [entry])
        log("Excel log is open in Excel right now - saved this entry to pending_log.csv; "
            "it will be merged into Download_Log.xlsx on the next run.")


def _append_csv(path: Path, entries):
    new = not path.exists()
    with open(path, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=COLUMNS)
        if new:
            w.writeheader()
        for e in entries:
            w.writerow(e)


def _read_pending(log_path: Path):
    p = _pending_path(log_path)
    if not p.exists():
        return []
    with open(p, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    for r in rows:
        for k in ("Duration (min)", "File Size (MB)"):
            try:
                r[k] = float(r[k]) if r.get(k) not in (None, "") else None
            except ValueError:
                r[k] = None
    return rows


def _write_workbook(log_path: Path, new_entries):
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    if log_path.exists():
        wb = load_workbook(log_path)
        ws = wb["Log"]
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Log"
        ws.append(COLUMNS)

    ok_fill = PatternFill("solid", fgColor="E6F4EA")
    bad_fill = PatternFill("solid", fgColor="FDE7E7")
    for e in new_entries:
        ws.append([e.get(c) for c in COLUMNS])
        status_cell = ws.cell(row=ws.max_row, column=COLUMNS.index("Status") + 1)
        status_cell.fill = ok_fill if e.get("Status") == "Success" else bad_fill

    _style_header(ws, WIDTHS)
    ws.auto_filter.ref = ws.dimensions
    for row in ws.iter_rows(min_row=2, min_col=COLUMNS.index("Error") + 1,
                            max_col=COLUMNS.index("Error") + 1):
        for c in row:
            c.alignment = Alignment(wrap_text=False)

    rows = [dict(zip(COLUMNS, r)) for r in ws.iter_rows(min_row=2, values_only=True)]
    _rebuild_summary(wb, rows)
    _rebuild_by_day(wb, rows)
    wb.save(log_path)


def _style_header(ws, widths):
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
    head_fill = PatternFill("solid", fgColor="2452A8")
    for i, cell in enumerate(ws[1], start=1):
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        if i <= len(widths):
            ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    ws.freeze_panes = "A2"


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _date_key(d):
    try:
        return datetime.strptime(str(d), "%d-%m-%Y")
    except ValueError:
        return datetime.min


def _rebuild_summary(wb, rows):
    if "Summary" in wb.sheetnames:
        del wb["Summary"]
    ws = wb.create_sheet("Summary", 0)
    head = ["Client", "Report", "Attempts", "Successful", "Failed", "Success Rate",
            "Last Successful Download", "Last File", "Avg Duration (min)", "Latest Size (MB)"]
    ws.append(head)
    groups = defaultdict(list)
    for r in rows:
        groups[(r["Client"] or "", r["Report"] or "")].append(r)
    for (client, report), items in sorted(groups.items(), key=lambda kv: (kv[0][0].lower(), kv[0][1])):
        ok = [r for r in items if r["Status"] == "Success"]
        durations = [_num(r["Duration (min)"]) for r in ok if _num(r["Duration (min)"]) is not None]
        last_ok = max(ok, key=lambda r: (_date_key(r["Date"]), str(r["Finished At"]))) if ok else None
        ws.append([
            client, report, len(items), len(ok), len(items) - len(ok),
            len(ok) / len(items) if items else 0,
            f"{last_ok['Date']} {last_ok['Finished At']}" if last_ok else "never",
            last_ok["File Name"] if last_ok else "",
            round(sum(durations) / len(durations), 1) if durations else None,
            _num(last_ok["File Size (MB)"]) if last_ok else None,
        ])
    for c in ws["F"][1:]:
        c.number_format = "0%"
    _style_header(ws, [22, 28, 10, 11, 8, 12, 24, 34, 18, 16])
    ws.cell(row=ws.max_row + 2, column=1,
            value=f"Last updated {datetime.now():%d-%m-%Y %H:%M}")


def _rebuild_by_day(wb, rows):
    if "By Day" in wb.sheetnames:
        del wb["By Day"]
    ws = wb.create_sheet("By Day", 1)
    ws.append(["Date", "Day", "Attempts", "Successful", "Failed", "Total MB Downloaded",
               "Clients", "Reports Downloaded"])
    days = defaultdict(list)
    for r in rows:
        days[r["Date"]].append(r)
    for d in sorted(days, key=_date_key, reverse=True):
        items = days[d]
        ok = [r for r in items if r["Status"] == "Success"]
        ws.append([
            d, items[0]["Day"], len(items), len(ok), len(items) - len(ok),
            round(sum(_num(r["File Size (MB)"]) or 0 for r in ok), 1),
            ", ".join(sorted({r["Client"] for r in items if r["Client"]})),
            ", ".join(sorted({r["Report"] for r in ok if r["Report"]})),
        ])
    _style_header(ws, [12, 6, 10, 11, 8, 20, 30, 60])
