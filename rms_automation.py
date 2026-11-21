"""
IntelliPrism (RMS) reports engine.

Reuses the proven sign-in + Silent Login steps from automation.py, then on the
client's IntelliPrism site, for each selected report, ONE AT A TIME (the site only
builds one report at a time):

  Intelliprism Reports -> card's GENERATE REPORT -> filter panel
  -> required date range = 01-Jan-2016 .. today (calendar icon, pick both dates, SET DATES)
  -> GENERATE REPORT -> "Report is generating" -> Downloads
  -> wait for THIS run's row (Requested on >= trigger time) to show Success
  -> DOWNLOAD (opens a tab that delivers the file) -> save -> next report.
"""
from __future__ import annotations

import re
import time
from datetime import date, datetime, timedelta
from pathlib import Path

from playwright.sync_api import TimeoutError as PWTimeoutError, sync_playwright

import automation as core

RMS_REPORTS = [
    "Request Profile Report",
    "Positions Report",
    "Candidate Evaluation Report",
    "IJP - Demand View Report",
    "IJP - Applicant View Report",
]
KNOWN_REPORTS = [("IntelliPrism Reports", RMS_REPORTS)]

START_DATE = date(2016, 1, 1)          # always from 1 Jan 2016 ...
READY_TIMEOUT_S = 30 * 60              # ... wait up to 30 min for a report to be built
POLL_S = 15
TRIGGER_TOLERANCE_S = 180              # "Requested on" is shown to the minute

_REQ_RE = re.compile(r"(\d{1,2} [A-Za-z]{3} \d{4}),?\s+(\d{1,2}:\d{2}\s*[AP]M)", re.IGNORECASE)


def _origin(page) -> str:
    return re.match(r"(https?://[^/]+)", page.url).group(1)


def _exact(text: str):
    return re.compile(rf"^\s*{re.escape(text)}\s*$", re.IGNORECASE)


# ------------------------------------------------------------------ navigation
# The site only opens these pages from its left menu; typing the address sends you
# back to Home. So we always click the menu, like a person does.
def _menu(page, label: str):
    return page.locator("a.menu-list-item").filter(has_text=_exact(label)).first


DEFAULT_ROLE = "RMG Lead"
ROLE = DEFAULT_ROLE   # set per run from the app's "Role" box


def _role_name(label: str) -> str:
    return re.sub(r"\(\s*my role\s*\)", "", label or "", flags=re.I).strip()


def _handle_role_screen(page, log) -> bool:
    """After Silent Login IntelliPrism may ask which role to continue as
    (an ng-select 'userRole' + an arrow button). Pick ROLE if offered."""
    sel = page.locator("ng-select[name='userRole']").first
    try:
        if sel.count() == 0 or not sel.is_visible():
            return False
    except Exception:
        return False
    current = _role_name(sel.locator(".ng-value-label").first.inner_text(timeout=3000))
    chosen = current
    try:
        sel.locator(".ng-select-container").first.click(timeout=5000)
        opts = page.locator("ng-dropdown-panel .ng-option")
        opts.first.wait_for(timeout=8000)
        labels = [_role_name(t) for t in opts.all_inner_texts()]
        log(f"Role screen - roles offered: {', '.join(labels)}")
        want = (ROLE or "").strip().lower()
        idx = next((i for i, l in enumerate(labels) if l.lower() == want), None)
        if idx is None:
            idx = next((i for i, l in enumerate(labels) if want and want in l.lower()), None)
        if idx is not None:
            opts.nth(idx).click(timeout=5000)
            chosen = labels[idx]
        else:
            log(f"Role '{ROLE}' isn't offered - continuing as '{current}'.")
            page.keyboard.press("Escape")
    except PWTimeoutError:
        log("Couldn't open the role list - continuing with the role already selected.")
    page.wait_for_timeout(500)
    nxt = page.locator(".next-btn").first
    try:
        nxt.click(timeout=8000)
    except Exception:
        nxt.evaluate("el => el.click()")
    log(f"Continuing as role '{chosen}'.")
    core.wait_networkidle(page, log, timeout=30000)
    return True


def _ensure_in_app(page, origin: str, log):
    if "/rms/" not in page.url or "file-download" in page.url:
        page.goto(f"{origin}/rms/home", timeout=60000)
        core.wait_networkidle(page, log, timeout=30000)
    if core.looks_like_login_page(page):
        raise core.SessionExpired("Redirected to login while opening IntelliPrism.")
    menu = _menu(page, "Intelliprism Reports")
    end = time.time() + 60
    while time.time() < end:
        try:
            if menu.count() and menu.is_visible():
                return
        except Exception:
            pass
        if _handle_role_screen(page, log):
            continue
        page.wait_for_timeout(1000)
    core.debug_dump(page, "rms_menu_not_found", log)
    raise core.ReportNotFound(
        f"IntelliPrism's left menu has no 'Intelliprism Reports' (role used: '{ROLE}'). "
        "Check the Role box - this role may not have reports access.")


def _click_menu(page, label: str, log):
    item = _menu(page, label)
    try:
        item.click(timeout=8000)
    except Exception:
        item.evaluate("el => el.click()")
    core.wait_networkidle(page, log, timeout=30000)


def go_to_rms_reports(page, log=core._noop_log, origin: str | None = None):
    origin = origin or _origin(page)
    log("Opening Intelliprism Reports...")
    _ensure_in_app(page, origin, log)
    for attempt in (1, 2):
        _click_menu(page, "Intelliprism Reports", log)
        try:
            page.locator("mat-card.report-card h3").first.wait_for(timeout=30000)
            page.wait_for_timeout(1500)
            return
        except PWTimeoutError:
            if attempt == 1:
                log("Report cards didn't show yet - clicking Home, then Intelliprism Reports again...")
                _click_menu(page, "Home", log)
    core.debug_dump(page, "rms_reports_not_loaded", log)
    raise core.ReportNotFound("Intelliprism Reports didn't load (no report cards). "
                              "Does this admin user have IntelliPrism access?")


def go_to_downloads(page, origin: str, log=core._noop_log):
    """Open (or refresh) Downloads via the menu. If we're already there, hop to
    Intelliprism Reports first so the table reloads with the latest statuses."""
    _ensure_in_app(page, origin, log)
    if "download-center" in page.url:
        _click_menu(page, "Intelliprism Reports", log)
        page.wait_for_timeout(800)
    _click_menu(page, "Downloads", log)
    try:
        page.locator("td.mat-column-name").first.wait_for(timeout=40000)
    except PWTimeoutError:
        pass


def _card(page, report_name: str):
    return page.locator("mat-card.report-card").filter(
        has=page.locator("h3").filter(has_text=_exact(report_name)))


def _drawer(page):
    return page.locator("mat-drawer.right-side-nav")


def _drawer_open(page) -> bool:
    try:
        d = _drawer(page)
        return d.count() > 0 and d.first.get_by_text("Select filters", exact=False).first.is_visible()
    except Exception:
        return False


# ------------------------------------------------------------------ steps
def open_filter_panel(page, report_name: str, log=core._noop_log, stop_event=None):
    card = _card(page, report_name)
    if card.count() == 0:
        visible = [t.strip() for t in page.locator("mat-card.report-card h3").all_inner_texts()]
        raise core.ReportNotFound(f"Report '{report_name}' not on Intelliprism Reports. Visible: {visible}")
    btn = card.first.locator("button").filter(has_text=_exact("GENERATE REPORT")).first

    # If another report is still being generated the buttons are disabled - wait it out.
    end = time.time() + 15 * 60
    while btn.is_disabled() and time.time() < end:
        core._check_stop(stop_event)
        log("Another report is still generating - waiting before starting this one...")
        page.wait_for_timeout(20000)
        go_to_rms_reports(page, log)
        btn = _card(page, report_name).first.locator("button").filter(has_text=_exact("GENERATE REPORT")).first

    for how in ("click", "direct", "force"):
        try:
            if how == "click":
                btn.click(timeout=5000)
            elif how == "direct":
                btn.evaluate("el => el.click()")
            else:
                btn.scroll_into_view_if_needed(timeout=3000)
                btn.click(force=True, timeout=5000)
        except Exception as exc:
            log(f"GENERATE REPORT {how} failed: {str(exc).splitlines()[0][:100]}")
        for _ in range(16):
            if _drawer_open(page):
                page.wait_for_timeout(800)
                return
            page.wait_for_timeout(500)
        log(f"Filter panel didn't open after {how} - trying again...")
    core.debug_dump(page, f"rms_filter_panel_not_opened_{report_name}", log)
    raise core.ReportNotFound(f"Filter panel for '{report_name}' didn't open.")


_MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August",
           "September", "October", "November", "December"]
_DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


def _fmt_field(d: date) -> str:
    return f"{d.day:02d}-{_MONTHS[d.month - 1][:3]}-{d.year}"   # e.g. 01-Jan-2016 (locale-independent)


def _pick_day(page, picker, d: date, log):
    picker.locator("select[aria-label='Select month']").select_option(str(d.month))
    picker.locator("select[aria-label='Select year']").select_option(str(d.year))
    page.wait_for_timeout(400)
    aria = f"{_DAYS[d.weekday()]}, {_MONTHS[d.month - 1]} {d.day}, {d.year}"   # "Friday, January 1, 2016"
    cell = picker.locator(f"div.ngb-dp-day[aria-label='{aria}']").first
    cell.wait_for(timeout=5000)
    target = cell.locator(".custom-day").first if cell.locator(".custom-day").count() else cell
    target.click(timeout=5000)
    page.wait_for_timeout(300)


def set_required_date_range(page, report_name: str, start: date, end: date, log=core._noop_log):
    """The first date field in every panel is the required one (marked *)."""
    drawer = _drawer(page).first
    field = drawer.locator("input[placeholder='dd-mm-yyyy']").first
    field.wait_for(timeout=10000)
    # Structure: div.date-filter > div.input-group > [input, div.calendar-btn (icon)] + calendar dropdown
    box = field.locator("xpath=ancestor::div[contains(@class,'date-filter')][1]")
    cal = field.locator("xpath=following-sibling::*[contains(@class,'calendar-btn')][1]")
    if cal.count() == 0:
        cal = box.locator(".calendar-btn").first
    try:
        cal.click(timeout=5000)
    except Exception:
        cal.click(force=True, timeout=5000)
    picker = box.locator("ngb-datepicker").first
    try:
        picker.wait_for(state="visible", timeout=8000)
    except PWTimeoutError:
        cal.evaluate("el => el.click()")
        picker.wait_for(state="visible", timeout=8000)

    _pick_day(page, picker, start, log)
    _pick_day(page, picker, end, log)
    set_btn = box.locator("button").filter(has_text=_exact("SET DATES")).first
    set_btn.click(timeout=5000)
    page.wait_for_timeout(800)

    want = f"{_fmt_field(start)} to {_fmt_field(end)}"
    got = (field.input_value() or "").strip()
    if got.lower() != want.lower():
        core.debug_dump(page, f"rms_date_not_set_{report_name}", log)
        raise core.ReportNotFound(f"Date range didn't stick for '{report_name}': expected '{want}', got '{got}'.")
    log(f"Date range set: {got}")


def submit_generate(page, report_name: str, log=core._noop_log) -> datetime:
    drawer = _drawer(page).first
    gen = drawer.locator("button").filter(has_text=_exact("GENERATE REPORT")).last
    queued_at = datetime.now()
    gen.click(timeout=10000)
    # "Report is generating" dialog -> GO TO DOWNLOADS (we navigate ourselves anyway)
    dialog = page.locator("mat-dialog-container").filter(has_text=re.compile("generating", re.I))
    try:
        dialog.first.wait_for(timeout=15000)
        log("Report generation started.")
        go = dialog.first.locator("button").filter(has_text=re.compile(r"go to downloads", re.I)).first
        if go.count():
            go.click(timeout=8000)  # same as the person does; also closes the pop-up
            core.wait_networkidle(page, log, timeout=30000)
    except PWTimeoutError:
        log("Didn't see the 'Report is generating' message - will check Downloads anyway.")
    except Exception as exc:
        log(f"Couldn't click GO TO DOWNLOADS ({str(exc).splitlines()[0][:80]}) - opening Downloads from the menu.")
    return queued_at


def _parse_requested(text: str):
    m = _REQ_RE.search(text or "")
    if not m:
        return None
    try:
        return datetime.strptime(f"{m.group(1)} {m.group(2).replace(' ', '')}", "%d %b %Y %I:%M%p")
    except ValueError:
        return None


def _find_this_run_row(page, report_name: str, not_before: datetime):
    """Newest row for this report requested at/after not_before. Returns (row, status) or (None, None)."""
    rows = page.locator("tbody tr")
    for i in range(min(rows.count(), 25)):
        row = rows.nth(i)
        try:
            name = row.locator("td.mat-column-name").inner_text(timeout=2000).strip()
            if name.lower() != report_name.lower():
                continue
            req = _parse_requested(row.locator("td.mat-column-requeston").inner_text(timeout=2000))
            if req is None or req < not_before.replace(second=0, microsecond=0):
                continue
            status = row.locator("td.mat-column-status").inner_text(timeout=2000).strip()
            return row, status
        except Exception:
            continue
    return None, None


def wait_and_download(page, origin: str, client_name: str, report_name: str, queued_at: datetime,
                      log=core._noop_log, stop_event=None) -> Path | None:
    not_before = queued_at - timedelta(seconds=TRIGGER_TOLERANCE_S)
    started = time.time()
    last_note = 0.0
    while time.time() - started < READY_TIMEOUT_S:
        core._check_stop(stop_event)
        go_to_downloads(page, origin, log)
        if core.looks_like_login_page(page):
            raise core.SessionExpired("Redirected to login while checking Downloads.")
        row, status = _find_this_run_row(page, report_name, not_before)
        if row is not None and re.search(r"success|completed|ready", status or "", re.I):
            return _download_row(page, row, client_name, report_name, log)
        if row is not None and re.search(r"fail|error", status or "", re.I):
            core.debug_dump(page, f"rms_report_failed_{report_name}", log)
            raise core.ReportNotFound(f"Downloads shows '{status}' for '{report_name}'.")
        if time.time() - last_note > 45:
            waited = int(time.time() - started)
            log(f"Waiting for '{report_name}' to finish ({status or 'not listed yet'}) - {waited // 60}:{waited % 60:02d} so far...")
            last_note = time.time()
        end = time.time() + POLL_S
        while time.time() < end:
            core._check_stop(stop_event)
            time.sleep(1)
    core.debug_dump(page, f"rms_not_ready_{report_name}", log)
    return None


def _download_row(page, row, client_name: str, report_name: str, log) -> Path:
    """DOWNLOAD opens a new tab that delivers the file; catch the download from any tab."""
    ctx = page.context
    got = []
    attached = set()

    def attach(p):
        if id(p) not in attached:
            attached.add(id(p))
            p.on("download", lambda d: got.append(d))

    for p in ctx.pages:
        attach(p)
    ctx.on("page", attach)
    btn = row.locator("td.mat-column-action button").first
    btn.click(timeout=10000)
    end = time.time() + 120
    while not got and time.time() < end:
        page.wait_for_timeout(500)
    if not got:
        core.debug_dump(page, f"rms_download_not_started_{report_name}", log)
        raise core.ReportNotFound(f"Clicked DOWNLOAD for '{report_name}' but no file arrived within 2 minutes.")
    dl = got[0]
    suggested = dl.suggested_filename or ""
    ext = Path(suggested).suffix or ".xlsx"
    now = datetime.now()
    day_dir = core.client_download_dir(client_name) / now.strftime("%d-%m-%Y")
    day_dir.mkdir(parents=True, exist_ok=True)
    save_path = day_dir / f"{core.sanitize(report_name)}_{now.strftime('%H%M%S')}{ext}"
    dl.save_as(str(save_path))
    for p in list(ctx.pages):  # close the helper tab(s) DOWNLOAD opened
        if p is not page and "file-download" in (p.url or ""):
            try:
                p.close()
            except Exception:
                pass
    log(f"Saved to {save_path}")
    return save_path


# ------------------------------------------------------------------ orchestration
def run_job(client_name: str, admin_user: str, report_names, log=print, headless=True,
            stop_event=None, notify_on_expiry=None, trigger="Manual", send_email=True, role=None):
    """Generate + download each report in turn. Same contract as automation.run_job."""
    import report_log

    global ROLE
    ROLE = (role or DEFAULT_ROLE).strip()
    core.EMAIL_PRODUCT = "IntelliPrism"

    report_names = list(report_names)
    if not core.has_valid_session():
        raise RuntimeError("Not signed in yet - click Sign in first.")
    run_id = datetime.now().strftime("%Y%m%d-%H%M%S")
    results = {"ok": [], "failed": []}
    details, pending = [], list(report_names)

    def record(report, status, started, file_path=None, error=None):
        finished = datetime.now()
        details.append({"report": report, "status": status, "file": str(file_path) if file_path else "",
                        "error": error, "minutes": round((finished - started).total_seconds() / 60, 1)})
        (results["ok"] if status == "Success" else results["failed"]).append(report)
        if report in pending:
            pending.remove(report)
        core._status(report, status, file_path, error)
        try:
            report_log.append_entry(core.REPORT_LOG_PATH, report_log.make_entry(
                run_id, trigger, client_name, admin_user, report, status, started, finished,
                file_path=file_path, error=error), log=log)
        except Exception as exc:
            log(f"Could not update the Excel log: {exc}")

    core.ensure_download_root(log)
    job_started = datetime.now()
    try:
        with sync_playwright() as p:
            try:
                context = p.chromium.launch_persistent_context(
                    str(core.USER_DATA_DIR), channel=core.BROWSER_CHANNEL, headless=headless,
                    accept_downloads=True, no_viewport=True,
                    args=["--start-maximized", "--test-type"] if not headless else None)
            except Exception as exc:
                raise RuntimeError(f"Could not open the signed-in {core.BROWSER_CHANNEL} profile - close any "
                                   f"other copy of the app first. ({exc})")
            page = context.pages[0] if context.pages else context.new_page()
            try:
                try:
                    core.select_client_and_silent_login(page, client_name, admin_user, log)
                except core.SessionExpired:
                    if notify_on_expiry:
                        core.send_session_expired_email(log)
                    raise
                origin = _origin(page)
                log(f"Logged in as '{admin_user}' for '{client_name}'. Opening Intelliprism Reports...")
                for n, report in enumerate(report_names, 1):
                    core._check_stop(stop_event)
                    rlog = lambda m, r=report: log(f"[{r}] {m}")
                    started = datetime.now()
                    log(f"REPORT {n}/{len(report_names)} {report}")
                    try:
                        go_to_rms_reports(page, log, origin)
                        open_filter_panel(page, report, rlog, stop_event)
                        set_required_date_range(page, report, START_DATE, date.today(), rlog)
                        queued_at = submit_generate(page, report, rlog)
                        core._status(report, "Generating")
                        log(f"WAITING {report}")
                        saved = wait_and_download(page, origin, client_name, report, queued_at, rlog, stop_event)
                        if saved:
                            record(report, "Success", started, file_path=saved)
                        else:
                            record(report, "Failed", started,
                                   error=f"Not ready within {READY_TIMEOUT_S // 60} minutes.")
                    except (core.SessionExpired, core.RunStopped):
                        raise
                    except Exception as exc:
                        err = str(exc).splitlines()[0][:300]
                        rlog(f"ERROR: {err}")
                        record(report, "Failed", started, error=err)
            finally:
                context.close()
    except Exception as exc:
        reason = ("Stopped by user." if isinstance(exc, core.RunStopped)
                  else "Login session expired." if isinstance(exc, core.SessionExpired)
                  else str(exc).splitlines()[0][:300])
        status = "Stopped" if isinstance(exc, core.RunStopped) else "Failed"
        for r in list(pending):
            record(r, status, job_started, error=reason)
        if send_email:
            core.send_run_summary_email(client_name, admin_user, trigger, details, log)
        raise
    if send_email:
        core.send_run_summary_email(client_name, admin_user, trigger, details, log)
    return results
