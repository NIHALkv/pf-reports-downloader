"""
Playwright automation core for the SkillPrism (Prismforce) Reports Center
downloader - Desktop edition.
"""
import os
import re
import smtplib
import time
from email.mime.text import MIMEText
from pathlib import Path
from datetime import datetime, timedelta

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError

import config as _cfg

GLOBAL_ADMIN_URL = "https://pf.prismforce.com/globalAdmin/clients"
BROWSER_CHANNEL = os.environ.get("BROWSER_CHANNEL", "msedge")  # "msedge" or "chrome"
# Where the browser profile, debug screenshots and run logs live:
# %APPDATA%\PF Reports Downloader for the .exe, this folder when run from source.
BASE_DIR = _cfg.DATA_DIR

try:
    CONFIG = _cfg.load_config()
    CONFIG_ERROR = None
except _cfg.ConfigError as exc:
    CONFIG = {
        "user_email": None, "alert_to_email": None, "gmail_app_password": None,
        "clients": [], "profile_dir_name": "unconfigured_profile", "expected_domain": "prismforce.ai",
    }
    CONFIG_ERROR = str(exc)

USER_DATA_DIR = BASE_DIR / f"{BROWSER_CHANNEL}_{CONFIG['profile_dir_name']}"
# Where reports are saved: <DOWNLOAD_ROOT>/<Client>/<DD-MM-YYYY>/<Report>_<HHMMSS>.xlsx
# Override with "download_root" in config.json if needed.
DOWNLOAD_ROOT = Path(CONFIG.get("download_root") or r"C:\Reports download")
REPORT_LOG_PATH = DOWNLOAD_ROOT / "Download_Log.xlsx"
DEBUG_DIR = BASE_DIR / "debug"

# Optional callback for a UI: STATUS_HOOK(report_name, status, file_path, error)
# status is one of "Generating", "Success", "Failed", "Stopped".
STATUS_HOOK = None


def _status(report, status, file_path=None, error=None):
    if STATUS_HOOK:
        try:
            STATUS_HOOK(report, status, file_path, error)
        except Exception:
            pass


def set_download_root(path) -> None:
    """Change where reports are saved (used by the desktop app's 'Save to' box)."""
    global DOWNLOAD_ROOT, REPORT_LOG_PATH
    DOWNLOAD_ROOT = Path(path)
    REPORT_LOG_PATH = DOWNLOAD_ROOT / "Download_Log.xlsx"


class FolderError(Exception):
    """The save folder doesn't exist and can't be created."""


def ensure_download_root(log=lambda _m: None) -> Path:
    """Create the save folder (and any parent folders) if it isn't there yet."""
    root = Path(DOWNLOAD_ROOT)
    if root.is_dir():
        return root
    try:
        root.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise FolderError(
            f"Couldn't create the save folder '{root}' ({exc.strerror or exc}). "
            "Pick another folder in 'Save to', e.g. one inside Documents.") from exc
    log(f"Created the save folder: {root}")
    return root


def set_post_trigger_wait(seconds: int) -> None:
    global POST_TRIGGER_WAIT_S, POST_GENERATE_DELAY_S
    POST_TRIGGER_WAIT_S = POST_GENERATE_DELAY_S = int(seconds)


# ---------------------------------------------------------------- email settings
EMAIL_PRODUCT = "SkillPrism"   # subject prefix; the IntelliPrism engine sets its own
SEND_ATTEMPT_EMAIL = True      # desktop apps turn this off: they only send the "completed" email
_KEYRING_SERVICE = "PF Report Downloader (Gmail app password)"


def get_app_password() -> str:
    """Gmail app password: Windows Credential Manager first, then config.json (legacy)."""
    user = CONFIG.get("user_email")
    pw = ""
    if user:
        try:
            import keyring

            pw = keyring.get_password(_KEYRING_SERVICE, user) or ""
        except Exception:
            pw = ""
    pw = pw or (CONFIG.get("gmail_app_password") or "")
    pw = pw.replace(" ", "")
    return "" if pw.startswith("REPLACE") else pw


def set_app_password(pw: str) -> str:
    """Store the Gmail app password securely. Returns where it was stored."""
    user = CONFIG.get("user_email")
    if not user:
        raise ValueError("Sign in first, so the app knows which Gmail account sends the email.")
    pw = (pw or "").replace(" ", "").strip()
    try:
        import keyring

        if pw:
            keyring.set_password(_KEYRING_SERVICE, user, pw)
        else:
            try:
                keyring.delete_password(_KEYRING_SERVICE, user)
            except Exception:
                pass
        return "Windows Credential Manager"
    except Exception:
        data = _cfg.load_config()
        data["gmail_app_password"] = pw
        _cfg.save_config(data)
        CONFIG["gmail_app_password"] = pw
        return "the app's settings file"


def emails_enabled() -> bool:
    return bool(get_app_password()) and bool(CONFIG.get("user_email"))


def send_test_email(to_addrs, log=lambda _m: None) -> bool:
    body = ("This is a test from the Prismforce Report Downloader.\n\n"
            "If you can read this, 'Email when done' is set up correctly: you'll get a summary "
            "here whenever a download finishes.")
    return _send_email(f"{EMAIL_PRODUCT}: test email from Report Downloader", body, to_addrs=to_addrs, log=log)

LOGIN_TIMEOUT_MS = 5 * 60 * 1000  # 5 minutes to complete Google SSO manually
REPORT_READY_TIMEOUT_S = 600
REPORT_POLL_INTERVAL_S = 8
# After ALL reports are triggered, wait this long before downloading (config "post_trigger_wait_s").
POST_TRIGGER_WAIT_S = int(CONFIG.get("post_trigger_wait_s") or 120)
POST_GENERATE_DELAY_S = POST_TRIGGER_WAIT_S  # old name, kept for compatibility
# A Download History row counts as "this run" if stamped no earlier than trigger time minus this.
TRIGGER_TS_TOLERANCE_S = 120

KNOWN_REPORTS = [
    ("Consolidated Reports", [
        "Employee Capabilities",
        "Employee Master",
        "Consolidated Activity Log",
        "Detailed Activity Log",
    ]),
    ("Skills", [
        "User Skill Data",
        "Skill Inventory Data",
        "Skill Master Data",
        "Skill Endorsement Data",
    ]),
    ("RSCs", [
        "User RSCs Data",
        "RSCs Inventory Data",
        "RSCs Master",
        "User RSC Skills Gap Report",
    ]),
    ("Designation Path", ["User Current Role Report"]),
    ("Learning Elements", ["User Learning Element Report"]),
]


class SessionExpired(Exception):
    """The saved login no longer works - re-run login()."""


class ReportNotFound(Exception):
    """A step couldn't find what it needed on the page, even after retrying."""


class RunStopped(Exception):
    """Cooperative cancellation - the user clicked Stop."""


def _noop_log(_msg):
    pass


def _check_stop(stop_event):
    if stop_event is not None and stop_event.is_set():
        raise RunStopped("Run stopped by user.")


def sanitize(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9._ -]+", "_", name).strip()


def client_download_dir(client_name: str) -> Path:
    """Reuses an existing client folder (matched case-insensitively, so
    'Acme' and 'ACME' land in the same place); creates one otherwise."""
    DOWNLOAD_ROOT.mkdir(parents=True, exist_ok=True)
    wanted = sanitize(client_name)
    for existing in DOWNLOAD_ROOT.iterdir():
        if existing.is_dir() and existing.name.strip().lower() == wanted.lower():
            return existing
    return DOWNLOAD_ROOT / wanted


def _exact_ci(text: str):
    return re.compile(f"^{re.escape(text)}$", re.IGNORECASE)


def has_valid_session() -> bool:
    return USER_DATA_DIR.exists() and any(USER_DATA_DIR.iterdir())


def looks_like_login_page(page) -> bool:
    return "Sign in with Google" in (page.content() or "") or page.url.startswith(
        "https://accounts.google.com"
    )


_EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")


def _detect_signed_in_email(page):
    try:
        text = page.content()
    except Exception:
        return None
    matches = _EMAIL_RE.findall(text or "")
    if not matches:
        return None
    domain = CONFIG.get("expected_domain")
    if domain:
        for m in matches:
            if m.lower().endswith("@" + domain):
                return m
    return matches[0]


def _try_auto_google_click(page, log=_noop_log) -> bool:
    domain = CONFIG.get("expected_domain") or "prismforce.ai"

    try:
        signin_btn = page.get_by_text("Sign in with Google", exact=False).first
        signin_btn.wait_for(timeout=8000)
        signin_btn.click()
        log("Clicked 'Sign in with Google'.")
    except PWTimeoutError:
        return False

    try:
        page.wait_for_url("**accounts.google.com**", timeout=8000)
    except PWTimeoutError:
        return False

    try:
        tile = page.locator(f'[data-identifier$="@{domain}"]').first
        tile.wait_for(timeout=8000)
        tile.click()
        log(f"Clicked your existing @{domain} account in Google's account chooser.")
        return True
    except PWTimeoutError:
        log(f"Google didn't offer an already-signed-in @{domain} account to click - "
            "it needs a password and/or 2FA, which this tool never enters. Please "
            "finish signing in by hand in the window.")
        return False


def _recover_via_google_reauth(page, log=_noop_log) -> bool:
    if not looks_like_login_page(page):
        return True
    log("Session looks expired - trying a silent Google re-auth first "
        "(existing cached account, no credentials entered)...")
    _try_auto_google_click(page, log)
    page.wait_for_timeout(1500)
    recovered = not looks_like_login_page(page)
    if recovered:
        log("Silent re-auth worked - no human needed.")
    return recovered


def login(log=_noop_log) -> bool:
    USER_DATA_DIR.mkdir(parents=True, exist_ok=True)
    if CONFIG_ERROR:
        log(f"Note: {CONFIG_ERROR} (continuing with a generic profile - the "
            "session-expired alert email won't work until config.json is set up.)")
    expected = CONFIG.get("user_email")
    log(f"Opening a {BROWSER_CHANNEL} window and attempting the sign-in clicks "
        f"automatically. If Google asks for a password, a code, or a device "
        f"confirmation, that part still needs you - waiting up to 5 minutes.")
    with sync_playwright() as p:
        try:
            context = p.chromium.launch_persistent_context(
                str(USER_DATA_DIR), channel=BROWSER_CHANNEL, headless=False,
                no_viewport=True, args=["--start-maximized", "--test-type"],
            )
        except Exception as exc:
            log(f"Could not open {BROWSER_CHANNEL} - is it installed? ({exc})")
            return False

        page = context.pages[0] if context.pages else context.new_page()
        page.goto("https://pf.prismforce.com/")
        _try_auto_google_click(page, log)
        try:
            page.wait_for_url(f"{GLOBAL_ADMIN_URL}*", timeout=LOGIN_TIMEOUT_MS)
        except Exception:
            page.goto(GLOBAL_ADMIN_URL)
            page.wait_for_timeout(3000)

        ok = "globalAdmin/clients" in page.url
        if ok:
            detected = _detect_signed_in_email(page)
            if detected:
                log(f"Signed in as {detected}.")
                if not CONFIG.get("user_email"):
                    try:  # remember it, so emails and the app header know who you are
                        CONFIG["user_email"] = detected
                        CONFIG["alert_to_email"] = CONFIG.get("alert_to_email") or detected
                        saved = _cfg.load_config() if _cfg.CONFIG_PATH.exists() else dict(_cfg.DEFAULTS)
                        saved["user_email"] = detected
                        saved["alert_to_email"] = saved.get("alert_to_email") or detected
                        _cfg.save_config(saved)
                    except Exception:
                        pass
                domain = CONFIG.get("expected_domain")
                if domain and not detected.lower().endswith("@" + domain):
                    log(f"WARNING: that doesn't look like a @{domain} account (per your "
                        f"config.json). Silent Login later may fail if this is the wrong "
                        f"account - re-run Login and pick the correct one if so.")
            log("Signed in - this profile will be reused on every run.")
        else:
            log(f"Still not on the Global Admin Clients page (current URL: {page.url}). "
                "Login was not captured - try again.")
        context.close()
        return ok


def _send_email(subject: str, body: str, to_addrs=None, log=_noop_log) -> bool:
    """Sends a plain-text email from your own Gmail account using the app password
    in config.json. Returns False (and logs why) instead of raising."""
    from_addr = CONFIG.get("user_email")
    app_password = get_app_password()
    to_addrs = to_addrs or [CONFIG.get("alert_to_email") or from_addr]
    to_addrs = [t for t in to_addrs if t]
    if not (from_addr and to_addrs and app_password):
        log("Email not sent: set your Gmail app password in the app ('Email when done' > Set app password).")
        return False
    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = from_addr
    msg["To"] = ", ".join(to_addrs)
    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=20) as server:
            server.starttls()
            server.login(from_addr, app_password)
            server.sendmail(from_addr, to_addrs, msg.as_string())
        log(f"Email sent to {', '.join(to_addrs)}: {subject}")
        return True
    except smtplib.SMTPAuthenticationError:
        log(f"Could not send email '{subject}': Gmail rejected the app password for {from_addr}. "
            "Create a new one and set it again in the app.")
        return False
    except Exception as exc:
        log(f"Could not send email '{subject}': {exc}")
        return False


def _notify_recipients():
    return CONFIG.get("notify_emails") or [CONFIG.get("alert_to_email") or CONFIG.get("user_email")]


def send_attempt_email(client_name, admin_user, trigger, report_names, queued, trigger_errors,
                       note=None, log=_noop_log) -> bool:
    """Email #1 - sent once every report has been triggered (or the run died before that).
    queued: {report: datetime queued}; trigger_errors: {report: error text}."""
    subject = f"{EMAIL_PRODUCT}: download attempted for {client_name} - {len(report_names)} report(s)"
    lines = [f"Client: {client_name}", f"Admin user: {admin_user}", f"Run type: {trigger}",
             f"Attempted: {datetime.now():%d-%m-%Y %H:%M}", "", "Reports attempted:"]
    for r in report_names:
        if r in queued:
            lines.append(f"  - {r}  (generation triggered at {queued[r]:%H:%M:%S})")
        elif r in trigger_errors:
            lines.append(f"  - {r}  (could NOT trigger: {trigger_errors[r]})")
        else:
            lines.append(f"  - {r}  (not reached)")
    lines.append("")
    if note:
        lines += [note, ""]
    elif queued:
        wait_txt = (f"{POST_TRIGGER_WAIT_S / 60:g} min" if POST_TRIGGER_WAIT_S >= 60
                    else f"{POST_TRIGGER_WAIT_S}s")
        lines += [f"Waiting {wait_txt} for the reports to be ready, then "
                  f"downloading. A second email will confirm success or failure.", ""]
    return _send_email(subject, "\n".join(lines), to_addrs=_notify_recipients(), log=log)


def send_run_summary_email(client_name, admin_user, trigger, details, log=_noop_log) -> bool:
    """Email #2 - final result. details: list of dicts from run_job
    (report, status, file, error, minutes)."""
    ok = [d for d in details if d["status"] == "Success"]
    bad = [d for d in details if d["status"] != "Success"]
    if ok and not bad:
        subject = f"{EMAIL_PRODUCT}: SUCCESS - all {len(ok)} report(s) downloaded for {client_name}"
    elif ok:
        subject = (f"{EMAIL_PRODUCT}: PARTIAL - {len(ok)} downloaded, {len(bad)} failed "
                   f"for {client_name}")
    else:
        subject = f"{EMAIL_PRODUCT}: FAILED - no reports downloaded for {client_name}"
    lines = [f"Client: {client_name}", f"Admin user: {admin_user}", f"Run type: {trigger}",
             f"Finished: {datetime.now():%d-%m-%Y %H:%M}", "",
             f"Reports attempted ({len(details)}):"]
    for d in details:
        mark = "OK  " if d["status"] == "Success" else "FAIL"
        lines.append(f"  [{mark}] {d['report']}")
    lines.append("")
    if ok:
        lines.append("Downloaded successfully:")
        for d in ok:
            size = ""
            try:
                size = f", {Path(d['file']).stat().st_size / (1024 * 1024):.1f} MB"
            except Exception:
                pass
            lines.append(f"  - {d['report']} ({d['minutes']} min{size})")
            lines.append(f"    {d['file']}")
        lines.append("")
    if bad:
        lines.append("Not downloaded:")
        for d in bad:
            lines.append(f"  - {d['report']}: {d['status']} - {d['error'] or 'see log'}")
        lines.append("")
    lines.append(f"Full history: {REPORT_LOG_PATH}")
    return _send_email(subject, "\n".join(lines), to_addrs=_notify_recipients(), log=log)


def send_session_expired_email(log=_noop_log) -> bool:
    from_addr = CONFIG.get("user_email")
    to_addr = CONFIG.get("alert_to_email") or from_addr
    app_password = CONFIG.get("gmail_app_password")
    if not (from_addr and to_addr and app_password):
        log("Skipping session-expired alert email: config.json is missing "
            "user_email / alert_to_email / gmail_app_password.")
        return False

    msg = MIMEText(
        "Your SkillPrism Reports Downloader automation ran on schedule and found "
        "that your saved Google sign-in no longer works.\n\n"
        "Please open the SkillPrism Reports Downloader app and click Login again "
        "(sign in with Google, then choose your prismforce.ai account) so the "
        "next scheduled run succeeds.\n\n"
        "This is an automated message - no action needed if you already know "
        "your session expired."
    )
    msg["Subject"] = "Action needed: SkillPrism automation login has expired"
    msg["From"] = from_addr
    msg["To"] = to_addr

    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=20) as server:
            server.starttls()
            server.login(from_addr, app_password)
            server.sendmail(from_addr, [to_addr], msg.as_string())
        log(f"Sent a 'please log in again' alert email to {to_addr}.")
        return True
    except Exception as exc:
        log(f"Could not send the session-expired alert email: {exc}")
        return False


def debug_dump(page, label: str, log=_noop_log):
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = DEBUG_DIR / f"{sanitize(label)}_{ts}.png"
    try:
        page.screenshot(path=str(path), full_page=True)
        log(f"Saved a debug screenshot to: {path}")
    except Exception as exc:
        log(f"Could not capture a debug screenshot: {exc}")
    try:  # [trial-diagnostics] also save the page HTML next to the screenshot
        path.with_suffix(".html").write_text(page.content(), encoding="utf-8")
    except Exception:
        pass
    return path


def wait_networkidle(page, log=_noop_log, timeout=30000):
    try:
        page.wait_for_load_state("networkidle", timeout=timeout)
    except PWTimeoutError:
        log(f"Page didn't reach 'networkidle' within {timeout}ms - continuing anyway.")


def list_visible_reports(page):
    names = []
    try:
        for btn in page.get_by_role("button", name="GENERATE REPORT", exact=True).all():
            try:
                card = btn.locator(
                    "xpath=ancestor::*[.//button[normalize-space()='GENERATE REPORT']][1]"
                )
                text = card.inner_text(timeout=2000)
                first_line = text.strip().splitlines()[0] if text.strip() else ""
                if first_line:
                    names.append(first_line)
            except Exception:
                continue
    except Exception:
        pass
    return names


def select_client_and_silent_login(page, client_name: str, admin_user: str, log=_noop_log):
    log(f"Opening Global Admin > Clients and searching for '{client_name}'...")
    try:
        page.goto(GLOBAL_ADMIN_URL, timeout=60000)
    except PWTimeoutError:
        page.goto(GLOBAL_ADMIN_URL, timeout=60000)
    wait_networkidle(page, log)

    if looks_like_login_page(page):
        if not _recover_via_google_reauth(page, log):
            raise SessionExpired("Redirected to login while opening Global Admin > Clients.")
        page.goto(GLOBAL_ADMIN_URL, timeout=60000)
        wait_networkidle(page, log)

    search_box = page.get_by_placeholder("Search Clients...")
    try:
        search_box.wait_for(timeout=20000)
    except PWTimeoutError:
        log("Client search box hasn't appeared yet - reloading and retrying...")
        page.reload()
        wait_networkidle(page, log)
        if looks_like_login_page(page):
            if not _recover_via_google_reauth(page, log):
                raise SessionExpired("Redirected to login while opening Global Admin > Clients.")
            page.goto(GLOBAL_ADMIN_URL, timeout=60000)
            wait_networkidle(page, log)
        search_box.wait_for(timeout=25000)
    search_box.fill(client_name)
    page.wait_for_timeout(500)

    client_cell = page.get_by_role("cell", name=_exact_ci(client_name)).first
    try:
        client_cell.wait_for(timeout=10000)
    except PWTimeoutError:
        debug_dump(page, f"client_not_found_{client_name}", log)
        raise ReportNotFound(f"Could not find client '{client_name}' in Global Admin > Clients.")
    client_cell.click()
    wait_networkidle(page, log)

    log(f"Searching for admin user '{admin_user}'...")
    user_search = page.get_by_placeholder("Search by name, employee ID or email")
    user_search.click()
    user_search.fill("")
    user_search.press_sequentially(admin_user, delay=80)

    suggestion = page.get_by_text(admin_user).first
    try:
        suggestion.wait_for(timeout=15000)
    except PWTimeoutError:
        log(f"'{admin_user}' not showing yet - still waiting (not retyping)...")
        try:
            suggestion.wait_for(timeout=25000)
        except PWTimeoutError:
            debug_dump(page, f"admin_user_not_found_{admin_user}", log)
            raise ReportNotFound(
                f"Could not find admin user '{admin_user}' for this client after 40s. "
                "Double check spelling and that this user has PF Admin / report access."
            )
    suggestion.click()
    page.wait_for_timeout(500)

    user_row = page.get_by_role("cell", name=_exact_ci(admin_user)).first
    user_row.click()

    log(f"Silent-logging in as '{admin_user}'...")
    silent_login_btn = page.get_by_role("button", name="Silent Login", exact=True)
    try:
        silent_login_btn.wait_for(timeout=10000)
    except PWTimeoutError:
        debug_dump(page, f"no_silent_login_button_{admin_user}", log)
        raise ReportNotFound(f"'{admin_user}' has no Silent Login button for this client.")
    silent_login_btn.click()

    try:
        page.wait_for_url(lambda url: "pf.prismforce.com" not in url, timeout=30000)
    except PWTimeoutError:
        pass
    wait_networkidle(page, log, timeout=30000)

    if "pf.prismforce.com" in page.url:
        debug_dump(page, "silent_login_did_not_redirect", log)
        raise ReportNotFound(
            f"Silent Login did not redirect away from pf.prismforce.com (still on {page.url})."
        )


def dismiss_prismsearch_tooltip(page, log=_noop_log):
    # The banner's close "x" is a <pf-icon icon="close"> inside .adoption-tooltip
    try:
        close_icon = page.locator(".adoption-tooltip pf-icon[icon='close'], .adoption-tooltip mat-icon:text-is('close')")
        if close_icon.count() and close_icon.first.is_visible():
            _robust_click(close_icon.first, timeout=1500)
            log("Dismissed the PrismSearch tooltip banner.")
            return
    except Exception:
        pass
    try:
        banner = page.get_by_text("faster using PrismSearch", exact=False).first
        if banner.count() == 0 or not banner.is_visible():
            return
        container = banner.locator("xpath=ancestor::*[self::div or self::section][1]")
        for candidate in (
            container.get_by_role("button"),
            container.locator("svg, [class*='close' i], [class*='dismiss' i]"),
        ):
            if candidate.count():
                candidate.first.click(timeout=2000)
                log("Dismissed the PrismSearch tooltip banner.")
                return
    except Exception:
        pass


def go_to_reports_center(page, log=_noop_log):
    origin = re.match(r"(https?://[^/]+)", page.url).group(1)
    log(f"Navigating to Reports Center on {origin}...")
    page.goto(f"{origin}/sms/dashboards-reports")
    wait_networkidle(page, log)
    try:
        page.get_by_text("Reports Center", exact=True).first.wait_for(timeout=20000)
        page.get_by_role("button", name="GENERATE REPORT", exact=True).first.wait_for(timeout=20000)
        _wait_cards_ready(page)
    except PWTimeoutError:
        debug_dump(page, "reports_center_slow_to_render", log)
        log("Reports Center took longer than expected to render - continuing anyway.")
    page.wait_for_timeout(2000)
    dismiss_prismsearch_tooltip(page, log)


def _wait_cards_ready(page, timeout_s=20):
    """Report cards first render as loading skeletons; clicks during that phase are
    ignored by the page. Wait until every card shows a real title, then a moment more."""
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        try:
            ready = page.evaluate("""() => {
                const cards = [...document.querySelectorAll('app-report-accordion')];
                if (!cards.length) return false;
                const busy = document.querySelector('[class*="skeleton" i]:not([skeletonloader])');
                return !busy && cards.every(c => (c.querySelector('h3')?.textContent || '').trim().length > 0);
            }""")
            if ready:
                break
        except Exception:
            pass
        page.wait_for_timeout(500)
    page.wait_for_timeout(1500)


def _robust_click(locator, timeout=3000):
    """Report cards live inside <mat-expansion-panel-header aria-disabled="true">,
    so Playwright's normal click() refuses anything inside the header forever.
    Try a normal click briefly, then a forced mouse click, then fire the click
    event directly on the element."""
    try:
        locator.click(timeout=timeout)
        return
    except PWTimeoutError:
        pass
    try:
        locator.click(timeout=timeout, force=True)
        return
    except PWTimeoutError:
        pass
    locator.dispatch_event("click")


def _go_home_then_reports(page, log=_noop_log):
    origin = re.match(r"(https?://[^/]+)", page.url).group(1)
    page.goto(f"{origin}/home")
    wait_networkidle(page, log)
    go_to_reports_center(page, log)


def find_report_card(page, report_name: str, log=_noop_log):
    heading = page.get_by_text(report_name, exact=True).first
    try:
        heading.wait_for(timeout=20000)
    except PWTimeoutError:
        log(f"'{report_name}' not found yet - reloading Reports Center and retrying...")
        page.reload()
        wait_networkidle(page, log)
        try:
            heading.wait_for(timeout=25000)
        except PWTimeoutError:
            visible = list_visible_reports(page)
            debug_dump(page, f"report_not_found_{report_name}", log)
            raise ReportNotFound(
                f"Could not find report '{report_name}' on Reports Center after a reload. "
                f"Reports visible instead: {visible if visible else 'none'}."
            )
    # Preferred: the whole <app-report-accordion> for this report (stable whether the
    # card is collapsed or expanded). Fallback: nearest ancestor with a GENERATE button.
    title_re = re.compile(rf"^\s*{re.escape(report_name)}\s*$", re.IGNORECASE)
    accordion = page.locator("app-report-accordion").filter(
        has=page.locator("h3").filter(has_text=title_re))
    if accordion.count():
        return accordion.first
    return heading.locator("xpath=ancestor::*[.//button[normalize-space()='GENERATE REPORT']][1]")


def collapse_card(page, card):
    """Best effort - closes an expanded card's filter panel. Never raises."""
    try:
        close_btn = card.get_by_role("button", name="Close", exact=True)
        if close_btn.count():
            _robust_click(close_btn.first, timeout=1500)
            page.wait_for_timeout(500)
    except Exception:
        pass


def trigger_report(page, card, report_name: str, log=_noop_log):
    """Queues one report. The card's GENERATE REPORT button first expands a filter
    panel; that panel's own GENERATE REPORT button actually queues the report.
    Returns the local time the report was queued."""
    dismiss_prismsearch_tooltip(page, log)
    # NOTE: the card header itself has role="button" and its accessible name *contains*
    # "GENERATE REPORT", so a loose get_by_role() match picks the header first and clicks
    # the middle of the card. Target the real <button> elements only.
    first_btn = card.locator(
        "mat-expansion-panel-header button.generate-report-btn, "
        "mat-expansion-panel-header button:text-is('GENERATE REPORT')").first
    try:
        first_btn.wait_for(timeout=15000)
    except PWTimeoutError:
        debug_dump(page, f"generate_button_missing_{report_name}", log)
        raise ReportNotFound(f"No GENERATE REPORT button found for '{report_name}'.")
    started_toast = page.get_by_text("Report generation has started", exact=False)
    panel_btn = card.locator(
        ".mat-expansion-panel-content button.generate-report-btn, "
        ".mat-expansion-panel-content button:text-is('GENERATE REPORT')")
    queued_at = datetime.now()

    def _opened(wait_ms):
        """True once the filter panel (or the 'generation started' toast) shows up."""
        end = time.time() + wait_ms / 1000
        while time.time() < end:
            try:
                if started_toast.count() and started_toast.first.is_visible():
                    return "toast"
                if panel_btn.count() and panel_btn.first.is_visible():
                    return "panel"
            except Exception:
                pass
            page.wait_for_timeout(300)
        return None

    def _scroll_and_click():
        first_btn.scroll_into_view_if_needed(timeout=3000)
        page.wait_for_timeout(300)
        first_btn.click(force=True, timeout=3000)

    attempts = [
        ("normal click", lambda: _robust_click(first_btn)),
        ("direct click", lambda: first_btn.evaluate("el => el.click()")),
        ("scroll + click", _scroll_and_click),
    ]
    state = None
    for round_no in (1, 2):
        for how, do_click in attempts:
            dismiss_prismsearch_tooltip(page, log)
            queued_at = datetime.now()
            try:
                do_click()
            except Exception as exc:
                log(f"GENERATE REPORT {how} failed: {str(exc).splitlines()[0][:120]}")
            state = _opened(6000)
            if state:
                break
            log(f"Filter panel didn't open after {how} - trying again...")
            page.wait_for_timeout(1500)
        if state or round_no == 2:
            break
        log("Still no filter panel - reloading Reports Center and retrying this report once...")
        go_to_reports_center(page, log)
        try:
            first_btn.wait_for(timeout=15000)
        except PWTimeoutError:
            break

    if state != "toast":
        if state != "panel":
            debug_dump(page, f"filter_panel_not_opened_{report_name}", log)
            raise ReportNotFound(
                f"Clicked GENERATE REPORT for '{report_name}' (3 ways, plus a page reload) but "
                "neither the filter panel nor the 'generation has started' message appeared.")
        queued_at = datetime.now()
        _robust_click(panel_btn.first)
        log("Filter panel opened - clicked its GENERATE REPORT to queue the report.")

    # Confirm it really queued: success toast, or the panel button turning disabled.
    deadline = time.time() + 15
    confirmed = False
    while time.time() < deadline:
        try:
            if started_toast.count() and started_toast.first.is_visible():
                confirmed = True
                break
            if panel_btn.count() and panel_btn.first.is_disabled():
                confirmed = True
                break
        except Exception:
            pass
        page.wait_for_timeout(500)
    if not confirmed:
        log("Didn't see a 'generation has started' confirmation - will still check "
            "Download History for it.")
    collapse_card(page, card)
    return queued_at


def open_download_history(page, card, log=_noop_log):
    """Opens the Download History dialog for a report card. The icon sits inside an
    aria-disabled panel header, so it's clicked via _robust_click and the dialog
    is confirmed open."""
    icon = card.locator(
        "[icon='download_2'], .download-history-btn, img[alt='download_2']"
    ).first
    try:
        icon.wait_for(timeout=15000)
    except PWTimeoutError:
        raise ReportNotFound("No Download History icon on this card yet.")
    dialog = page.locator("mat-dialog-container, .cdk-overlay-pane").filter(
        has_text="Download History").first
    for target in (icon.locator("mat-icon").first, icon):
        try:
            if not target.count():
                continue
            _robust_click(target, timeout=2000)
            dialog.wait_for(timeout=8000)
            return
        except PWTimeoutError:
            continue
    debug_dump(page, "download_history_icon_not_enabled", log)
    raise ReportNotFound("Clicked the Download History icon but the dialog did not open.")


def get_download_history_modal(page, log=_noop_log):
    dialog = page.locator("mat-dialog-container, .cdk-overlay-pane").filter(
        has_text="Download History"
    ).first
    try:
        dialog.wait_for(timeout=15000)
    except PWTimeoutError:
        debug_dump(page, "download_history_modal_not_visible", log)
        raise ReportNotFound("The Download History dialog did not appear within 15s.")
    return dialog


def close_modal(page, modal):
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(300)
        if modal.count() and modal.first.is_visible():
            _robust_click(modal.locator("mat-icon.close-icon, button").first, timeout=2000)
    except Exception:
        pass


_ROW_TIMESTAMP_RE = re.compile(r"\d{1,2} [A-Za-z]{3} \d{4} \d{2}:\d{2}:\d{2}")


def _parse_row_timestamp(row_text: str):
    m = _ROW_TIMESTAMP_RE.search(row_text)
    if not m:
        return None
    try:
        return datetime.strptime(m.group(0), "%d %b %Y %H:%M:%S")
    except ValueError:
        return None


def get_latest_history_row(modal):
    status_badges = modal.locator("xpath=.//*[not(*)]").filter(
        has_text=re.compile(r"ready|in progress|failed", re.IGNORECASE)
    )
    best = (None, None, None)
    for i in range(status_badges.count()):
        badge = status_badges.nth(i)
        row, text, ts = None, None, None
        for level in range(1, 8):
            candidate = badge.locator(f"xpath=ancestor::*[{level}]")
            try:
                candidate_text = candidate.inner_text(timeout=1000)
            except PWTimeoutError:
                break
            candidate_ts = _parse_row_timestamp(candidate_text)
            if candidate_ts:
                row, text, ts = candidate, candidate_text, candidate_ts
                break
        if row is None:
            continue
        if ts and (best[2] is None or ts > best[2]):
            best = (row, text, ts)
    return best


def _row_download_control(row):
    control = row.locator("mat-icon").filter(has_text="download")
    if control.count() == 0:
        control = row.locator("button")
    return control


def _is_this_run(row_ts, not_before):
    return row_ts is not None and row_ts >= not_before


def wait_for_report_ready(page, card, report_name, not_before, log=_noop_log,
                          timeout_s=REPORT_READY_TIMEOUT_S, poll_s=REPORT_POLL_INTERVAL_S,
                          stop_event=None):
    """Polls Download History until the row generated by THIS run (timestamp at or
    after not_before) has a download icon. Returns True / False."""
    start = time.time()
    deadline = start + timeout_s
    while time.time() < deadline:
        _check_stop(stop_event)
        modal = get_download_history_modal(page, log)
        row, row_text, row_ts = get_latest_history_row(modal)
        for _ in range(10):  # rows load asynchronously after the dialog opens
            if row_ts is not None:
                break
            page.wait_for_timeout(1000)
            row, row_text, row_ts = get_latest_history_row(modal)
        elapsed = int(time.time() - start)
        if _is_this_run(row_ts, not_before):
            if _row_download_control(row).count() > 0:
                log(f"Report (generated {row_ts:%H:%M:%S}) is ready.")
                return True
            if "failed" in row_text.lower():
                log(f"Report generation FAILED on the server after {elapsed}s.")
                debug_dump(page, f"report_generation_failed_{report_name}", log)
                close_modal(page, modal)
                return False
            log(f"Still generating... ({elapsed}s, checking again in {poll_s}s)")
        elif row_ts is not None:
            log(f"Newest row is from {row_ts:%d %b %H:%M:%S}, older than this run's trigger "
                f"({elapsed}s, checking again in {poll_s}s)")
        else:
            log(f"Download History has no rows yet ({elapsed}s, checking again in {poll_s}s)")

        for _ in range(poll_s):
            _check_stop(stop_event)
            time.sleep(1)
        close_modal(page, modal)
        _go_home_then_reports(page, log)
        card = find_report_card(page, report_name, log=log)
        open_download_history(page, card, log=log)
    log(f"Timed out after {timeout_s}s waiting for the report to be ready.")
    debug_dump(page, f"report_ready_timeout_{report_name}", log)
    return False


def download_ready_report(page, save_path: Path, not_before, log=_noop_log):
    modal = get_download_history_modal(page, log)
    row, row_text, row_ts = get_latest_history_row(modal)
    download_control = _row_download_control(row) if row is not None else None
    if row is None or not _is_this_run(row_ts, not_before) or download_control.count() == 0:
        raise ReportNotFound(
            "The latest Download History row was no longer this run's ready report by the "
            "time we went to download it."
        )

    download = None
    last_exc = None
    for attempt in range(1, 4):
        try:
            with page.expect_download(timeout=60000) as dl_info:
                _robust_click(download_control.last, timeout=5000)
            download = dl_info.value
            break
        except PWTimeoutError as exc:
            last_exc = exc
            log(f"Download didn't start within 60s (possibly a transient storage-side 503) - "
                f"attempt {attempt}/3.")
            if attempt == 1:
                debug_dump(page, "download_click_timeout", log)
            if attempt < 3:
                page.wait_for_timeout(3000)
    if download is None:
        raise ReportNotFound(
            "The download never started after 3 attempts, each timing out."
        ) from last_exc

    save_path.parent.mkdir(parents=True, exist_ok=True)
    download.save_as(str(save_path))
    close_modal(page, modal)
    return save_path


def fetch_report(page, client_name: str, report_name: str, queued_at, log=_noop_log,
                 stop_event=None):
    """Phase 2 for one report: open its Download History, wait for this run's row,
    download it. Returns the saved path, or None if it never became ready."""
    _check_stop(stop_event)
    not_before = queued_at - timedelta(seconds=TRIGGER_TS_TOLERANCE_S)
    _go_home_then_reports(page, log)  # fresh page so the card shows its latest state
    card = find_report_card(page, report_name, log=log)
    try:
        open_download_history(page, card, log=log)
    except ReportNotFound as exc:
        log(f"Download History didn't open ({exc}) - refreshing and retrying once...")
        page.wait_for_timeout(5000)
        _go_home_then_reports(page, log)
        card = find_report_card(page, report_name, log=log)
        open_download_history(page, card, log=log)

    if not wait_for_report_ready(page, card, report_name, not_before, log=log,
                                 stop_event=stop_event):
        return None

    now = datetime.now()
    day_dir = client_download_dir(client_name) / now.strftime("%d-%m-%Y")
    save_path = day_dir / f"{sanitize(report_name)}_{now.strftime('%H%M%S')}.xlsx"
    download_ready_report(page, save_path, not_before, log=log)
    log(f"Saved to {save_path}")
    return save_path


def _sleep_with_stop(seconds, stop_event, log=_noop_log, label="Waiting"):
    end = time.time() + seconds
    last_note = 0.0
    while True:
        _check_stop(stop_event)
        left = int(end - time.time())
        if left <= 0:
            return
        if time.time() - last_note >= 30:
            log(f"{label}: {left}s left...")
            last_note = time.time()
        time.sleep(1)


def run_job(client_name: str, admin_user: str, report_names, log=print,
            headless=True, stop_event=None, notify_on_expiry=None, trigger="Manual",
            send_email=True):
    """Runs one client's reports in two phases:
      1. Trigger GENERATE REPORT for every report (then email #1: 'attempted').
      2. Wait POST_TRIGGER_WAIT_S (default 2 min), then download each one.
    Every report is written to the Excel log, and email #2 (success/failure with
    the list of reports attempted) is sent at the end."""
    import report_log

    if notify_on_expiry is None:
        notify_on_expiry = headless
    report_names = list(report_names)
    if not has_valid_session():
        raise RuntimeError("Not signed in yet - click Login first.")

    run_id = datetime.now().strftime("%Y%m%d-%H%M%S")
    results = {"ok": [], "failed": []}
    details = []
    pending = list(report_names)
    queued = {}           # report -> datetime queued
    trigger_errors = {}   # report -> error text
    started_at = {}       # report -> datetime its attempt started
    state = {"attempt_email_sent": False}

    def record(report, status, started, file_path=None, error=None):
        finished = datetime.now()
        details.append({
            "report": report, "status": status, "file": str(file_path) if file_path else "",
            "error": error, "minutes": round((finished - started).total_seconds() / 60, 1),
        })
        (results["ok"] if status == "Success" else results["failed"]).append(report)
        _status(report, status, file_path, error)
        if report in pending:
            pending.remove(report)
        try:
            report_log.append_entry(REPORT_LOG_PATH, report_log.make_entry(
                run_id, trigger, client_name, admin_user, report, status, started, finished,
                file_path=file_path, error=error), log=log)
        except Exception as exc:
            log(f"Could not update the Excel log: {exc}")

    def attempt_email(note=None):
        if send_email and SEND_ATTEMPT_EMAIL and not state["attempt_email_sent"]:
            state["attempt_email_sent"] = True
            send_attempt_email(client_name, admin_user, trigger, report_names, queued,
                               trigger_errors, note=note, log=log)

    ensure_download_root(log)
    job_started = datetime.now()
    try:
        with sync_playwright() as p:
            try:
                context = p.chromium.launch_persistent_context(
                    str(USER_DATA_DIR), channel=BROWSER_CHANNEL, headless=headless,
                    accept_downloads=True, no_viewport=True,
                    args=["--start-maximized", "--test-type"] if not headless else None,
                )
            except Exception as exc:
                raise RuntimeError(
                    f"Could not open the signed-in {BROWSER_CHANNEL} profile - make sure "
                    f"{BROWSER_CHANNEL} isn't already open elsewhere with this same profile. "
                    f"Original error: {exc}"
                )
            page = context.pages[0] if context.pages else context.new_page()
            try:
                try:
                    select_client_and_silent_login(page, client_name, admin_user, log)
                except SessionExpired:
                    if notify_on_expiry:
                        send_session_expired_email(log)
                    raise
                log(f"Logged in as '{admin_user}' for '{client_name}'. Opening Reports Center...")
                go_to_reports_center(page, log)

                # ---- Phase 1: trigger every report ----
                log(f"PHASE 1 - triggering {len(report_names)} report(s)...")
                for report_name in report_names:
                    _check_stop(stop_event)
                    rlog = lambda m, r=report_name: log(f"[{r}] {m}")
                    started_at[report_name] = datetime.now()
                    try:
                        card = find_report_card(page, report_name, log=rlog)
                        queued[report_name] = trigger_report(page, card, report_name, log=rlog)
                        rlog(f"Triggered at {queued[report_name]:%H:%M:%S}.")
                        _status(report_name, "Generating")
                    except (SessionExpired, RunStopped):
                        raise
                    except Exception as exc:
                        err = str(exc).splitlines()[0][:300]
                        rlog(f"ERROR while triggering: {err}")
                        if not isinstance(exc, ReportNotFound):
                            debug_dump(page, f"trigger_error_{report_name}", log)
                        trigger_errors[report_name] = err
                        record(report_name, "Failed", started_at[report_name],
                               error=f"Could not trigger: {err}")
                        try:  # get back to a clean Reports Center for the next one
                            go_to_reports_center(page, log)
                        except Exception:
                            pass

                attempt_email(None if queued else "None of the reports could be triggered.")

                # ---- Phase 2: wait, then download each one ----
                if queued:
                    log(f"PHASE 2 - all triggered. Waiting {POST_TRIGGER_WAIT_S}s for the "
                        f"reports to be generated...")
                    _sleep_with_stop(POST_TRIGGER_WAIT_S, stop_event, log, "Waiting")
                    log("PHASE 3 - downloading...")
                for report_name, queued_at in queued.items():
                    _check_stop(stop_event)
                    rlog = lambda m, r=report_name: log(f"[{r}] {m}")
                    started = started_at[report_name]
                    try:
                        saved = fetch_report(page, client_name, report_name, queued_at,
                                             log=rlog, stop_event=stop_event)
                        if saved:
                            record(report_name, "Success", started, file_path=saved)
                        else:
                            record(report_name, "Failed", started,
                                   error="Report was not ready within the wait time.")
                    except SessionExpired:
                        record(report_name, "Failed", started, error="Login session expired.")
                        if notify_on_expiry:
                            send_session_expired_email(log)
                        raise
                    except RunStopped:
                        raise
                    except Exception as exc:
                        err = str(exc).splitlines()[0][:300]
                        rlog(f"ERROR while downloading: {err}")
                        if not isinstance(exc, ReportNotFound):
                            debug_dump(page, f"download_error_{report_name}", log)
                        record(report_name, "Failed", started, error=err)
            finally:
                context.close()
    except Exception as exc:
        # Anything that stopped the whole job: log the reports it never finished.
        reason = ("Stopped by user." if isinstance(exc, RunStopped)
                  else "Login session expired." if isinstance(exc, SessionExpired)
                  else str(exc).splitlines()[0][:300])
        status = "Stopped" if isinstance(exc, RunStopped) else "Failed"
        for r in list(pending):
            record(r, status, started_at.get(r, job_started), error=reason)
        attempt_email(f"Run ended early: {reason}")
        if send_email:
            send_run_summary_email(client_name, admin_user, trigger, details, log)
        raise
    if send_email:
        order = {r: i for i, r in enumerate(report_names)}
        details.sort(key=lambda d: order.get(d["report"], 999))
        send_run_summary_email(client_name, admin_user, trigger, details, log)
    return results
