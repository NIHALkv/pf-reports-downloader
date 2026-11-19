"""
Per-user settings for the SkillPrism Report Downloader.

Desktop app (.exe): settings live in %APPDATA%\\PF Reports Downloader\\config.json
and are created/updated by the app itself, so nobody edits a file by hand.
Running from source: config.json next to this file (same as before).

Nothing is required up front:
- user_email is filled in automatically after the first Sign in.
- gmail_app_password is optional; without it, summary emails are simply skipped.
"""
import json
import os
import re
import sys
from pathlib import Path

APP_NAME = "PF Reports Downloader"  # settings/sign-in folder name: kept unchanged so existing sign-ins carry over
BASE_DIR = Path(__file__).resolve().parent


def _data_dir() -> Path:
    if getattr(sys, "frozen", False):  # packaged .exe
        d = Path(os.environ.get("APPDATA", Path.home())) / APP_NAME
    else:
        d = BASE_DIR
    d.mkdir(parents=True, exist_ok=True)
    return d


DATA_DIR = _data_dir()
CONFIG_PATH = DATA_DIR / "config.json"
EXAMPLE_PATH = BASE_DIR / "config.example.json"

DEFAULTS = {
    "user_email": None,
    "alert_to_email": None,
    "gmail_app_password": "",
    "notify_emails": [],
    "clients": [],
    "download_root": r"C:\Reports download",
    "post_trigger_wait_s": 120,
    "expected_domain": "prismforce.ai",
    # remembered by the desktop app
    "recent_clients": [],      # [{"client_name": ..., "admin_user": ...}]
    "selected_reports": [],
    "show_browser": True,
}


class ConfigError(Exception):
    """config.json is malformed."""


def _sanitize_for_dirname(email: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", (email or "").strip().lower())


def load_config() -> dict:
    data = dict(DEFAULTS)
    if CONFIG_PATH.exists():
        try:
            data.update(json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
        except json.JSONDecodeError as exc:
            raise ConfigError(f"config.json is not valid JSON: {exc}") from exc
    data.setdefault("alert_to_email", data.get("user_email"))
    if not data.get("alert_to_email"):
        data["alert_to_email"] = data.get("user_email")
    # One fixed browser profile per Windows user, so a changed email never loses the sign-in.
    data["profile_dir_name"] = "profile" if getattr(sys, "frozen", False) else (
        f"{_sanitize_for_dirname(data['user_email'])}_profile" if data.get("user_email") else "profile")
    if data.get("user_email"):
        data["expected_domain"] = data["user_email"].split("@")[-1].lower()
    return data


def save_config(data: dict) -> None:
    """Persist settings, dropping values computed at load time."""
    keep = {k: v for k, v in data.items() if k not in ("profile_dir_name",)}
    CONFIG_PATH.write_text(json.dumps(keep, indent=2), encoding="utf-8")
