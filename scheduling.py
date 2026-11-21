"""
Scheduled downloads via Windows Task Scheduler.

Each schedule is saved in the app settings and registered as a Windows task that
starts the same .exe with  --scheduled <id>  at the chosen time/days. In that mode
the app shows no window: it runs the job in the background (reusing the saved
sign-in), writes a log, stores the result as "last run", and exits.
Runs even when the app is closed, as long as the PC is on and you're logged in.
"""
from __future__ import annotations

import subprocess
import sys
import uuid
from datetime import datetime
from pathlib import Path

import config as cfgmod

DAYS = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
PRODUCT_TASK_NAME = {
    "skillprism": "PF SkillPrism Report Downloader",
    "intelliprism": "PF IntelliPrism Report Downloader",
}


class ScheduleError(Exception):
    pass


def supported() -> bool:
    return sys.platform.startswith("win")


def task_name(s: dict) -> str:
    client = "".join(c for c in s["client"] if c not in '\\/:*?"<>|').strip()[:40]
    return f"{PRODUCT_TASK_NAME[s['product']]} - {client} - {s['id']}"


def _app_command(product: str, sched_id: str) -> str:
    """Command Windows runs: the .exe itself, or pythonw + entry script from source."""
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}" --scheduled {sched_id}'
    entry = Path(__file__).resolve().parent / f"app_{product}.py"
    py = Path(sys.executable)
    pyw = py.with_name("pythonw.exe") if py.with_name("pythonw.exe").exists() else py
    return f'"{pyw}" "{entry}" --scheduled {sched_id}'


def _schtasks(args: list[str]) -> subprocess.CompletedProcess:
    flags = 0x08000000 if supported() else 0  # CREATE_NO_WINDOW: no console flash
    return subprocess.run(["schtasks", *args], capture_output=True, text=True, creationflags=flags)


# ------------------------------------------------------------------ settings storage
def list_schedules(product: str) -> list[dict]:
    return [s for s in cfgmod.load_config().get("schedules", []) if s.get("product") == product]


def _save_all(update) -> None:
    data = cfgmod.load_config()
    data["schedules"] = update(list(data.get("schedules", [])))
    cfgmod.save_config(data)


def describe(s: dict) -> str:
    days = s["days"]
    when = "Every day" if len(days) == 7 else ("Weekdays" if days == DAYS[:5] else ", ".join(d.title() for d in days))
    return f"{when} at {s['time']}"


# ------------------------------------------------------------------ create / remove
def add_schedule(product: str, client: str, admin: str, reports: list[str], folder: str,
                 time_hhmm: str, days: list[str], role: str | None = None) -> dict:
    if not supported():
        raise ScheduleError("Scheduling uses Windows Task Scheduler, so it only works on Windows.")
    if not days:
        raise ScheduleError("Pick at least one day.")
    try:
        datetime.strptime(time_hhmm, "%H:%M")
    except ValueError:
        raise ScheduleError("Time must look like 08:30 (24-hour).") from None
    s = {
        "id": uuid.uuid4().hex[:6], "product": product, "client": client, "admin": admin,
        "role": role, "reports": list(reports), "folder": folder, "time": time_hhmm,
        "days": [d for d in DAYS if d in days], "created": datetime.now().isoformat(timespec="minutes"),
        "last_run": None,
    }
    sc = ["/SC", "DAILY"] if len(s["days"]) == 7 else ["/SC", "WEEKLY", "/D", ",".join(s["days"])]
    r = _schtasks(["/Create", "/TN", task_name(s), "/TR", _app_command(product, s["id"]), *sc,
                   "/ST", time_hhmm, "/F"])
    if r.returncode != 0:
        raise ScheduleError(f"Windows didn't accept the schedule: {(r.stderr or r.stdout).strip()[:200]}")
    _save_all(lambda all_: all_ + [s])
    return s


def remove_schedule(sched_id: str) -> None:
    target = next((s for s in cfgmod.load_config().get("schedules", []) if s["id"] == sched_id), None)
    if target and supported():
        _schtasks(["/Delete", "/TN", task_name(target), "/F"])  # fine if already gone
    _save_all(lambda all_: [s for s in all_ if s["id"] != sched_id])


def task_exists(s: dict) -> bool:
    return supported() and _schtasks(["/Query", "/TN", task_name(s)]).returncode == 0


# ------------------------------------------------------------------ the background run
def _record_last_run(sched_id: str, result: dict) -> None:
    def upd(all_):
        for s in all_:
            if s["id"] == sched_id:
                s["last_run"] = result
        return all_
    _save_all(upd)


def run_scheduled(product: str, sched_id: str) -> int:
    """Entry point for  app.exe --scheduled <id>  (no window)."""
    if sys.stdout is None or sys.stderr is None:  # windowed .exe: no console
        out = open(cfgmod.DATA_DIR / "app_output.txt", "a", encoding="utf-8", buffering=1)  # noqa: SIM115
        sys.stdout = sys.stdout or out
        sys.stderr = sys.stderr or out
    log_dir = cfgmod.DATA_DIR / "run_logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"scheduled_{datetime.now():%Y%m%d}.log"

    def log(msg: str) -> None:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n")

    s = next((x for x in list_schedules(product) if x["id"] == sched_id), None)
    if not s:
        log(f"Schedule {sched_id} not found in settings - nothing to do.")
        return 1
    log(f"=== Scheduled run '{describe(s)}' for {s['client']} ({len(s['reports'])} report(s)) ===")
    started = datetime.now()
    import automation as core

    core.set_download_root(s["folder"])
    kwargs = {"headless": True, "trigger": "Scheduled", "notify_on_expiry": True,
              "send_email": core.emails_enabled()}
    if product == "intelliprism":
        import rms_automation as engine
        kwargs["role"] = s.get("role")
    else:
        engine = core
    result = {"when": started.strftime("%d-%m-%Y %H:%M"), "overall": "FAILED", "ok": 0,
              "failed": len(s["reports"]), "error": None}
    try:
        res = engine.run_job(s["client"], s["admin"], s["reports"], log=log, **kwargs)
        ok, bad = len(res.get("ok", [])), len(res.get("failed", []))
        result.update(ok=ok, failed=bad, overall="SUCCESS" if ok and not bad else ("PARTIAL" if ok else "FAILED"))
    except core.SessionExpired:
        result["error"] = "Sign-in expired - open the app and click 'Sign in again'."
    except Exception as exc:  # noqa: BLE001
        result["error"] = str(exc).splitlines()[0][:200]
    log(f"Result: {result['overall']} ({result['ok']} ok, {result['failed']} failed)"
        + (f" - {result['error']}" if result["error"] else ""))
    _record_last_run(sched_id, result)
    return 0 if result["overall"] == "SUCCESS" else 2
