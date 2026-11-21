"""
SkillPrism Report Downloader: desktop app (this is what SkillPrism_Report_Downloader.exe runs).

Same proven engine as the SkillPrism Reports Downloader (automation.py), in a
branded window:
  Step 1  Sign in (Google SSO, remembered)
  Step 2  Client name + admin user (Silent Login)
  Step 3  Pick reports by category
  Then    Download reports: live steps, per-report status, Stop, Open folder.
"""
from __future__ import annotations

import os
import queue
import sys
import threading
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox, ttk

import config as cfgmod

# ---------------------------------------------------------------- brand tokens (PF design kit)
NAVY = "#001453"          # on-primary-container
PRIMARY = "#1f40cb"       # primary
PRIMARY_DARK = "#0933c1"  # on-primary-fixed-variant
ON_PRIMARY = "#ffffff"
SURFACE = "#f7f9ff"       # surface
CARD = "#ffffff"          # surface-container-lowest
TINT = "#ecf4ff"          # surface-container-low
TEXT = "#191c20"          # on-surface
MUTED = "#43474e"         # on-surface-variant
LINE = "#c3c7cf"          # outline-variant
OK = "#16a34a"            # status only
WARN = "#b45309"
BAD = "#dc2626"

STEPS = [("login", "Silent login"), ("generate", "Generate"), ("wait", "Wait"), ("download", "Download")]

# The same window runs two tools; each .exe picks one.
PRODUCTS = {
    "skillprism": {
        "name": "SkillPrism Report Downloader",
        "engine": "automation",
        "hint": "Generates all ticked reports, waits 2 minutes, then downloads them.",
    },
    "intelliprism": {
        "name": "IntelliPrism Report Downloader",
        "engine": "rms_automation",
        "hint": "One report at a time (the site's rule): date range 1 Jan 2016 to today, generate, wait, download.",
    },
}


def resource(rel: str) -> Path:
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)) / rel


def load_brand_font() -> None:
    ttf = resource("branding/PlusJakartaSans.ttf")
    if sys.platform.startswith("win") and ttf.exists():
        try:
            import ctypes

            ctypes.windll.gdi32.AddFontResourceExW(str(ttf), 0x10, 0)  # FR_PRIVATE
        except Exception:
            pass


class App:
    def __init__(self, root: tk.Tk, product: str = "skillprism") -> None:
        self.product = product
        self.P = PRODUCTS[product]
        self.root = root
        self.q: queue.Queue = queue.Queue()
        self.settings = cfgmod.load_config()
        self.busy = None            # None | "login" | "run"
        self.stop_event = None
        self.wait_until = 0.0
        self.wait_since = 0.0
        self.wait_label = ""
        self.log_file = cfgmod.DATA_DIR / "run_logs" / f"app_{datetime.now():%Y%m%d}.log"

        import automation  # engine (imports Playwright)

        self.engine = automation
        automation.STATUS_HOOK = lambda r, s, f=None, e=None: self.q.put(("status", (r, s, f, e)))
        if self.P["engine"] == "rms_automation":
            import rms_automation

            self.runner = rms_automation.run_job
            self.known_reports = rms_automation.KNOWN_REPORTS
        else:
            self.runner = automation.run_job
            self.known_reports = automation.KNOWN_REPORTS
        self.sel_key = "selected_reports" if product == "skillprism" else f"selected_reports_{product}"
        automation.EMAIL_PRODUCT = "IntelliPrism" if product == "intelliprism" else "SkillPrism"

        root.title(f"Prismforce · {self.P['name']}")
        root.configure(bg=SURFACE)
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        root.geometry(f"{min(1240, sw - 40)}x{min(800, sh - 80)}+10+10")
        root.minsize(min(900, sw - 40), min(560, sh - 80))
        if sys.platform.startswith("win"):
            try:
                root.state("zoomed")
            except tk.TclError:
                pass
        try:
            root.iconbitmap(default=str(resource("branding/pf-icon.ico")))
        except Exception:
            try:
                self._icon = tk.PhotoImage(file=str(resource("branding/pf-mark.png")))
                root.iconphoto(True, self._icon)
            except Exception:
                pass

        fams = set(tkfont.families(root))
        face = next((f for f in ("Plus Jakarta Sans", "Segoe UI", "Helvetica") if f in fams), "TkDefaultFont")
        self.F = {
            "title": (face, 15, "bold"), "h": (face, 11, "bold"), "body": (face, 10),
            "small": (face, 9), "btn": (face, 11, "bold"), "mono": ("Consolas", 9),
        }
        self._styles(face)
        self._header()
        self._footer()  # packed before the body so it always stays visible at the bottom
        main = tk.Frame(root, bg=SURFACE, padx=16, pady=10)
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=3, uniform="c")
        main.columnconfigure(1, weight=2, uniform="c")
        main.rowconfigure(0, weight=1)
        left = self._scrollable(main)
        left.master.master.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        right = tk.Frame(main, bg=SURFACE)
        right.grid(row=0, column=1, sticky="nsew")
        self._left(left)
        self._right(right)
        self._refresh_signin()
        root.protocol("WM_DELETE_WINDOW", self._on_close)
        root.after(150, self._poll)
        if os.environ.get("PF_SELFTEST"):  # packaging check used by the Windows build
            root.after(1500, self._selftest)

    # ------------------------------------------------------------ styles / building blocks
    def _styles(self, face: str) -> None:
        st = ttk.Style(self.root)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure("PF.Horizontal.TProgressbar", troughcolor=TINT, background=PRIMARY,
                     bordercolor=TINT, lightcolor=PRIMARY, darkcolor=PRIMARY, thickness=8)
        st.configure("PF.Treeview", background=CARD, fieldbackground=CARD, foreground=TEXT,
                     rowheight=26, font=(face, 10), borderwidth=0)
        st.configure("PF.Treeview.Heading", background=TINT, foreground=MUTED, font=(face, 9, "bold"),
                     relief="flat", borderwidth=0)
        st.map("PF.Treeview", background=[("selected", TINT)], foreground=[("selected", TEXT)])
        st.configure("PF.TCheckbutton", background=CARD, foreground=TEXT, font=(face, 9))
        st.map("PF.TCheckbutton", background=[("active", CARD)])
        st.configure("PF.TCombobox", padding=5)

    def _header(self) -> None:
        bar = tk.Frame(self.root, bg=NAVY, height=60)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        try:
            img = tk.PhotoImage(file=str(resource("branding/logo-on-dark.png")))
            self._logo = img.subsample(3, 3)  # keeps aspect ratio
            tk.Label(bar, image=self._logo, bg=NAVY).pack(side="left", padx=20)
        except Exception:
            tk.Label(bar, text="Prismforce", bg=NAVY, fg=ON_PRIMARY, font=self.F["h"]).pack(side="left", padx=20)
        tk.Label(bar, text=self.P["name"], bg=NAVY, fg=ON_PRIMARY, font=self.F["h"]).pack(side="right", padx=20)

    def _footer(self) -> None:
        """Action bar pinned to the bottom of the window: always visible, whatever the screen size."""
        bar = tk.Frame(self.root, bg=CARD, highlightthickness=1, highlightbackground=LINE)
        bar.pack(side="bottom", fill="x")
        inner = tk.Frame(bar, bg=CARD, padx=16, pady=10)
        inner.pack(fill="x")
        self.run_btn = self._button(inner, "Download reports", self.do_run, big=True)
        self.run_btn.pack(side="left")
        self.stop_btn = self._button(inner, "Stop", self.do_stop, kind="danger", big=True)
        self.stop_btn.pack(side="left", padx=10)
        self._set_enabled(self.stop_btn, False)
        self._button(inner, "Open folder", self._open_folder, kind="secondary", big=True).pack(side="left")
        self._button(inner, "Open logs", self._open_logs, kind="secondary", big=True).pack(side="left", padx=10)
        # run result (All done / Partly done / errors): always visible here
        self.result = tk.Label(inner, text="", bg=CARD, font=self.F["h"], anchor="w", justify="left")
        self.result.pack(side="left", padx=(16, 0), fill="x", expand=True)

    def _scrollable(self, parent) -> tk.Frame:
        """A frame that scrolls vertically when the screen is too small for all of it."""
        holder = tk.Frame(parent, bg=SURFACE)
        canvas = tk.Canvas(holder, bg=SURFACE, highlightthickness=0, bd=0)
        sb = ttk.Scrollbar(holder, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=SURFACE)
        win = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)

        def refresh(_e=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            need = inner.winfo_reqheight() > canvas.winfo_height() + 2
            if need and not sb.winfo_ismapped():
                sb.pack(side="right", fill="y")
            elif not need and sb.winfo_ismapped():
                sb.pack_forget()
                canvas.yview_moveto(0)

        inner.bind("<Configure>", refresh)
        canvas.bind("<Configure>", lambda e: (canvas.itemconfigure(win, width=e.width), refresh()))
        canvas.pack(side="left", fill="both", expand=True)

        def wheel(e):
            if sb.winfo_ismapped():
                step = -1 if (getattr(e, "delta", 0) > 0 or getattr(e, "num", 0) == 4) else 1
                canvas.yview_scroll(step * 3, "units")

        for w in (canvas, inner):
            w.bind("<Enter>", lambda e: (canvas.bind_all("<MouseWheel>", wheel),
                                         canvas.bind_all("<Button-4>", wheel), canvas.bind_all("<Button-5>", wheel)))
            w.bind("<Leave>", lambda e: (canvas.unbind_all("<MouseWheel>"),
                                         canvas.unbind_all("<Button-4>"), canvas.unbind_all("<Button-5>")))
        return inner

    def _card(self, parent, title: str | None = None, expand=False) -> tk.Frame:
        outer = tk.Frame(parent, bg=LINE)
        outer.pack(fill="both" if expand else "x", expand=expand, pady=(0, 10))
        inner = tk.Frame(outer, bg=CARD, padx=14, pady=10)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        if title:
            tk.Label(inner, text=title, bg=CARD, fg=PRIMARY, font=self.F["h"], anchor="w").pack(fill="x", pady=(0, 6))
        body = tk.Frame(inner, bg=CARD)  # content goes here (may use grid)
        body.pack(fill="both", expand=True)
        return body

    def _button(self, parent, text, cmd, kind="primary", big=False) -> tk.Button:
        bg, fg, hover = {
            "primary": (PRIMARY, ON_PRIMARY, PRIMARY_DARK),
            "secondary": (CARD, PRIMARY, TINT),
            "danger": (CARD, BAD, TINT),
            "chip": (TINT, TEXT, LINE),
        }[kind]
        b = tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg, activebackground=hover, activeforeground=fg,
                      font=self.F["btn"] if big else self.F["body"], relief="flat", bd=0, cursor="hand2",
                      padx=18 if big else 10, pady=8 if big else 3,
                      highlightthickness=0 if kind in ("primary", "chip") else 1, highlightbackground=LINE,
                      disabledforeground=MUTED)
        b._pf = (bg, hover)
        b.bind("<Enter>", lambda e: b.config(bg=hover) if str(b["state"]) != "disabled" else None)
        b.bind("<Leave>", lambda e: b.config(bg=bg) if str(b["state"]) != "disabled" else None)
        return b

    def _set_enabled(self, b: tk.Button, on: bool) -> None:
        bg, _ = b._pf
        b.config(state="normal" if on else "disabled", bg=bg if on else (LINE if bg == PRIMARY else CARD),
                 cursor="hand2" if on else "arrow")

    # ------------------------------------------------------------ left column: steps 1-3
    def _left(self, parent) -> None:
        # Step 1
        c1 = self._card(parent, "Step 1: Sign in")
        row = tk.Frame(c1, bg=CARD)
        row.pack(fill="x")
        self.login_btn = self._button(row, "Sign in", self.do_login)
        self.login_btn.pack(side="left")
        self.signin_lbl = tk.Label(row, text="", bg=CARD, font=self.F["body"], anchor="w", justify="left")
        self.signin_lbl.pack(side="left", padx=12)

        # Step 2
        c2 = self._card(parent, "Step 2: Client & admin user")
        c2.columnconfigure(1, weight=1)
        recent = self.settings.get("recent_clients") or []
        tk.Label(c2, text="Client Name", bg=CARD, fg=TEXT, font=self.F["body"]).grid(row=0, column=0, sticky="w", pady=4)
        self.client = ttk.Combobox(c2, values=[r["client_name"] for r in recent], font=self.F["body"])
        self.client.grid(row=0, column=1, sticky="we", padx=(12, 0), pady=4)
        self.client.bind("<<ComboboxSelected>>", self._client_picked)
        tk.Label(c2, text="Admin User", bg=CARD, fg=TEXT, font=self.F["body"]).grid(row=1, column=0, sticky="w", pady=4)
        self.admin = ttk.Combobox(c2, values=sorted({r["admin_user"] for r in recent}), font=self.F["body"])
        self.admin.grid(row=1, column=1, sticky="we", padx=(12, 0), pady=4)
        hint_row = 2
        self.role = None
        if self.product == "intelliprism":
            tk.Label(c2, text="Role", bg=CARD, fg=TEXT, font=self.F["body"]).grid(row=2, column=0, sticky="w", pady=4)
            self.role = ttk.Combobox(c2, values=["RMG Lead", "Admin", "Hiring Manager"], font=self.F["body"])
            self.role.set(self.settings.get("rms_role") or "RMG Lead")
            self.role.grid(row=2, column=1, sticky="we", padx=(12, 0), pady=4)
            hint_row = 3
        tk.Label(c2, text='Exactly as in Global Admin. The admin user must have "Silent Login" access for this client.'
                 + (" Role: the one picked on IntelliPrism's role screen." if self.product == "intelliprism" else ""),
                 bg=CARD, fg=MUTED, font=self.F["small"], wraplength=560, justify="left").grid(
                 row=hint_row, column=0, columnspan=2, sticky="w", pady=(2, 0))
        if recent:
            self.client.set(recent[0]["client_name"])
            self.admin.set(recent[0]["admin_user"])

        # Step 3
        c3 = self._card(parent, "Step 3: Reports to download")
        tools = tk.Frame(c3, bg=CARD)
        tools.pack(fill="x", pady=(0, 8))
        self._button(tools, "Select all", lambda: self._select_all(True), kind="chip").pack(side="left")
        self._button(tools, "Clear all", lambda: self._select_all(False), kind="chip").pack(side="left", padx=8)
        self.count_lbl = tk.Label(tools, text="", bg=CARD, fg=MUTED, font=self.F["small"])
        self.count_lbl.pack(side="left", padx=4)
        tk.Label(c3, text=self.P["hint"], bg=CARD, fg=MUTED, font=self.F["small"], anchor="w",
                 justify="left").pack(fill="x", pady=(0, 6))
        grid = tk.Frame(c3, bg=CARD)
        grid.pack(fill="x")
        chosen = set(self.settings.get(self.sel_key) or [])
        self.report_vars: dict[str, tk.BooleanVar] = {}
        single = len(self.known_reports) == 1
        for i, (group, names) in enumerate(self.known_reports):
            box = tk.Frame(grid, bg=LINE)
            if single:
                box.grid(row=0, column=0, columnspan=3, sticky="nsew", pady=(0, 8))
            else:
                box.grid(row=i // 3, column=i % 3, sticky="nsew", padx=(0, 8), pady=(0, 8))
            inner = tk.Frame(box, bg=CARD, padx=10, pady=8)
            inner.pack(fill="both", expand=True, padx=1, pady=1)
            tk.Label(inner, text=group, bg=CARD, fg=PRIMARY, font=self.F["small"] + ("bold",)).grid(
                row=0, column=0, columnspan=2, sticky="w")
            for j, n in enumerate(names):
                v = tk.BooleanVar(value=n in chosen)
                v.trace_add("write", lambda *a: self._update_count())
                self.report_vars[n] = v
                cb = ttk.Checkbutton(inner, text=n, variable=v, style="PF.TCheckbutton")
                if single:
                    cb.grid(row=1 + j // 2, column=j % 2, sticky="w", pady=1, padx=(0, 24))
                else:
                    cb.grid(row=1 + j, column=0, sticky="w", pady=1)
        for col in range(3):
            grid.columnconfigure(col, weight=1, uniform="g")
        self._update_count()

        # Options + actions
        c4 = self._card(parent)
        c4.columnconfigure(1, weight=1)
        tk.Label(c4, text="Save to", bg=CARD, fg=TEXT, font=self.F["body"]).grid(row=0, column=0, sticky="w")
        self.folder = tk.Entry(c4, font=self.F["body"], relief="solid", bd=1, fg=TEXT, highlightthickness=0)
        self.folder.insert(0, self.settings.get("download_root") or r"C:\Reports download")
        self.folder.grid(row=0, column=1, sticky="we", padx=(12, 8), ipady=3)
        self._button(c4, "Browse…", self._browse, kind="secondary").grid(row=0, column=2)
        self.show_browser = tk.BooleanVar(value=bool(self.settings.get("show_browser", True)))
        ttk.Checkbutton(c4, text="Show the browser while it works", variable=self.show_browser,
                        style="PF.TCheckbutton").grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))
        tk.Label(c4, text="The folder is created automatically if it doesn't exist.", bg=CARD, fg=MUTED,
                 font=self.F["small"]).grid(row=2, column=0, columnspan=3, sticky="w", pady=(2, 0))

        self._email_card(parent)
        self._schedule_card(parent)


    # ------------------------------------------------------------ email when done
    def _email_card(self, parent) -> None:
        self.engine.SEND_ATTEMPT_EMAIL = False  # the apps only send the "completed" email
        c = self._card(parent, "Email when done (optional)")
        c.columnconfigure(1, weight=1)
        self.notify_on = tk.BooleanVar(value=bool(self.settings.get("notify_on", False)))
        ttk.Checkbutton(c, text="Email a summary when the download finishes", variable=self.notify_on,
                        style="PF.TCheckbutton", command=self._email_state).grid(row=0, column=0, columnspan=3,
                                                                                 sticky="w")
        tk.Label(c, text="To", bg=CARD, fg=TEXT, font=self.F["body"]).grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.notify_to = tk.Entry(c, font=self.F["body"], relief="solid", bd=1, fg=TEXT, highlightthickness=0)
        default_to = self.settings.get("notify_to") or self.engine.CONFIG.get("user_email") or ""
        self.notify_to.insert(0, default_to)
        self.notify_to.grid(row=1, column=1, columnspan=2, sticky="we", padx=(12, 0), pady=(6, 0), ipady=3)
        tk.Label(c, text="One or more addresses, separated by commas.", bg=CARD, fg=MUTED,
                 font=self.F["small"]).grid(row=2, column=1, columnspan=2, sticky="w", padx=(12, 0))
        row = tk.Frame(c, bg=CARD)
        row.grid(row=3, column=0, columnspan=3, sticky="we", pady=(8, 0))
        self.pw_lbl = tk.Label(row, text="", bg=CARD, font=self.F["small"], anchor="w", justify="left")
        self.pw_lbl.pack(side="left")
        self._button(row, "Send test email", self._test_email, kind="chip").pack(side="right")
        self._button(row, "Set app password…", self._set_app_password, kind="chip").pack(side="right", padx=8)
        self._email_state()

    def _email_state(self) -> None:
        who = self.engine.CONFIG.get("user_email")
        if not who:
            txt, col = "Sent from your Gmail. Sign in first (Step 1).", MUTED
        elif self.engine.get_app_password():
            txt, col = f"Sent from {who}  ·  Gmail app password saved ✓", OK
        else:
            txt, col = f"Sent from {who}  ·  needs a Gmail app password (one-time setup)", WARN
        self.pw_lbl.config(text=txt, fg=col)
        self.notify_to.config(state="normal" if self.notify_on.get() else "disabled", disabledbackground=TINT)

    def _recipients(self) -> list[str]:
        import re as _re

        return [a.strip() for a in _re.split(r"[,;\s]+", self.notify_to.get()) if a.strip()]

    def _valid_recipients(self) -> list[str] | None:
        import re as _re

        to = self._recipients()
        bad = [a for a in to if not _re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", a)]
        if not to or bad:
            messagebox.showwarning("Email when done", "Enter valid email address(es) in 'To'"
                                   + (f" (not valid: {', '.join(bad)})." if bad else "."))
            return None
        return to

    def _set_app_password(self) -> None:
        if not self.engine.CONFIG.get("user_email"):
            messagebox.showinfo("App password", "Sign in first (Step 1), so the app knows which Gmail sends the email.")
            return
        if messagebox.askyesno(
                "Gmail app password",
                "Emails are sent from your own Gmail using a Google 'app password' (16 letters).\n\n"
                "1. On the Google page that opens, create an app password (name it e.g. 'Report Downloader').\n"
                "2. Copy it and paste it in the next box.\n\n"
                "It's stored in Windows Credential Manager on this PC only.\n\nOpen the Google page now?"):
            import webbrowser

            webbrowser.open("https://myaccount.google.com/apppasswords")
        from tkinter import simpledialog

        pw = simpledialog.askstring("Gmail app password", "Paste the 16-letter app password:", show="•",
                                    parent=self.root)
        if pw is None:
            return
        try:
            where = self.engine.set_app_password(pw)
        except ValueError as exc:
            messagebox.showwarning("App password", str(exc))
            return
        self._email_state()
        if pw.strip():
            messagebox.showinfo("App password", f"Saved in {where}. Click 'Send test email' to check it works.")

    def _test_email(self) -> None:
        to = self._valid_recipients()
        if not to:
            return
        if not self.engine.emails_enabled():
            self._set_app_password()
            if not self.engine.emails_enabled():
                return
        msgs = []
        ok = self.engine.send_test_email(to, log=msgs.append)
        (messagebox.showinfo if ok else messagebox.showerror)(
            "Test email", ("Sent! Check the inbox of " + ", ".join(to)) if ok else (msgs[-1] if msgs else "Not sent."))

    def _email_choice(self) -> tuple[bool, list[str]]:
        """(send?, recipients) for a run, after validating. Returns (False, []) when off."""
        if not self.notify_on.get():
            return False, []
        to = self._valid_recipients()
        if to is None:
            raise ValueError("bad recipients")
        if not self.engine.emails_enabled():
            if not messagebox.askyesno("Email when done", "No Gmail app password is set, so no email can be sent.\n\n"
                                       "Continue without email? (Choose No to set it up first.)"):
                raise ValueError("setup first")
            return False, []
        return True, to

    # ------------------------------------------------------------ schedules
    def _schedule_card(self, parent) -> None:
        import scheduling

        self.sched = scheduling
        c = self._card(parent, "Schedule automatic downloads")
        row = tk.Frame(c, bg=CARD)
        row.pack(fill="x")
        tk.Label(row, text="Time", bg=CARD, fg=TEXT, font=self.F["body"]).pack(side="left")
        self.s_hour = ttk.Spinbox(row, from_=0, to=23, width=3, format="%02.0f", wrap=True, font=self.F["body"])
        self.s_hour.set("08")
        self.s_hour.pack(side="left", padx=(8, 2))
        tk.Label(row, text=":", bg=CARD, fg=TEXT, font=self.F["body"]).pack(side="left")
        self.s_min = ttk.Spinbox(row, from_=0, to=55, increment=5, width=3, format="%02.0f", wrap=True,
                                 font=self.F["body"])
        self.s_min.set("30")
        self.s_min.pack(side="left", padx=(2, 16))
        self.s_days = {}
        for d in scheduling.DAYS:
            v = tk.BooleanVar(value=d not in ("SAT", "SUN"))
            self.s_days[d] = v
            ttk.Checkbutton(row, text=d.title(), variable=v, style="PF.TCheckbutton").pack(side="left", padx=(0, 4))

        row2 = tk.Frame(c, bg=CARD)
        row2.pack(fill="x", pady=(8, 6))
        self.s_add = self._button(row2, "Add schedule", self._add_schedule, kind="secondary")
        self.s_add.pack(side="left")
        self.s_remove = self._button(row2, "Remove selected", self._remove_schedule, kind="danger")
        self.s_remove.pack(side="left", padx=8)
        tk.Label(row2, text="Uses the client, admin user, reports and folder chosen above.", bg=CARD, fg=MUTED,
                 font=self.F["small"]).pack(side="left", padx=6)

        self.s_tree = ttk.Treeview(c, columns=("client", "reports", "last"), show="tree headings",
                                   style="PF.Treeview", height=3)
        self.s_tree.heading("#0", text="WHEN", anchor="w")
        self.s_tree.heading("client", text="CLIENT", anchor="w")
        self.s_tree.heading("reports", text="REPORTS", anchor="w")
        self.s_tree.heading("last", text="LAST RUN", anchor="w")
        self.s_tree.column("#0", width=170)
        self.s_tree.column("client", width=130)
        self.s_tree.column("reports", width=70)
        self.s_tree.column("last", width=190)
        for tag, col in (("ok", OK), ("bad", BAD), ("warn", WARN), ("idle", MUTED)):
            self.s_tree.tag_configure(tag, foreground=col)
        self.s_tree.pack(fill="x")
        note = ("Runs in the background even when this app is closed, as long as the PC is on and you're "
                "logged in to Windows. Each run's log is in Open logs > run_logs.")
        if not scheduling.supported():
            note = "Scheduling uses Windows Task Scheduler and is only available on Windows."
            self._set_enabled(self.s_add, False)
        tk.Label(c, text=note, bg=CARD, fg=MUTED, font=self.F["small"], wraplength=600, justify="left",
                 anchor="w").pack(fill="x", pady=(4, 0))
        self._refresh_schedules()
        self.root.after(60000, self._schedule_tick)

    def _refresh_schedules(self) -> None:
        try:
            items = self.sched.list_schedules(self.product)
        except Exception:
            items = []
        self.s_tree.delete(*self.s_tree.get_children())
        for s in items:
            lr = s.get("last_run")
            if lr:
                last = f"{lr['when']}  {lr['overall'].title()}" + (f" ({lr['ok']}/{lr['ok'] + lr['failed']})"
                                                                   if lr.get("ok") is not None else "")
                if lr.get("error"):
                    last += f" - {lr['error'][:40]}"
                tag = {"SUCCESS": "ok", "PARTIAL": "warn"}.get(lr["overall"], "bad")
            else:
                last, tag = "Not run yet", "idle"
            client = s["client"] + ("  ✉" if s.get("notify") else "")
            self.s_tree.insert("", "end", iid=s["id"], text=self.sched.describe(s),
                               values=(client, len(s["reports"]), last), tags=(tag,))

    def _schedule_tick(self) -> None:
        """Every minute, show results of background runs without reopening the app."""
        if not self.s_tree.selection():
            self._refresh_schedules()
        self.root.after(60000, self._schedule_tick)

    def _add_schedule(self) -> None:
        client, admin = self.client.get().strip(), self.admin.get().strip()
        reports = [n for n, v in self.report_vars.items() if v.get()]
        folder = self.folder.get().strip()
        if not client or not admin or not reports or not folder:
            messagebox.showwarning("Schedule", "Fill in Client Name, Admin User, at least one report and "
                                               "'Save to' first. The schedule uses these.")
            return
        try:
            hh, mm = int(self.s_hour.get()), int(self.s_min.get())
            assert 0 <= hh <= 23 and 0 <= mm <= 59
        except Exception:
            messagebox.showwarning("Schedule", "Enter a valid time, e.g. 08 : 30.")
            return
        days = [d for d, v in self.s_days.items() if v.get()]
        role = self.role.get().strip() if self.role is not None else None
        notify = []
        if self.notify_on.get():
            notify = self._valid_recipients()
            if notify is None:
                return
            if not self.engine.emails_enabled() and not messagebox.askyesno(
                    "Email when done", "No Gmail app password is set, so this schedule can't email yet.\n\n"
                    "Add the schedule anyway? (You can set the password later.)"):
                return
        try:
            s = self.sched.add_schedule(self.product, client, admin, reports, folder,
                                        f"{hh:02d}:{mm:02d}", days, role=role, notify=notify)
        except self.sched.ScheduleError as exc:
            messagebox.showerror("Schedule", str(exc))
            return
        self._save_settings(client, admin, reports)
        self._refresh_schedules()
        messagebox.showinfo("Schedule added",
                            f"{self.sched.describe(s)}: {len(reports)} report(s) for {client} will download "
                            f"automatically to {folder}."
                            + (f"\nA summary will be emailed to {', '.join(notify)}." if notify else "")
                            + "\n\nKeep the PC on and stay logged in to Windows at "
                            "that time. If your sign-in expires, the run will tell you to sign in again here.")

    def _remove_schedule(self) -> None:
        sel = self.s_tree.selection()
        if not sel:
            messagebox.showinfo("Schedule", "Click a schedule in the list first, then Remove selected.")
            return
        if not messagebox.askyesno("Remove schedule", "Remove the selected schedule?"):
            return
        for sid in sel:
            self.sched.remove_schedule(sid)
        self._refresh_schedules()

    # ------------------------------------------------------------ right column: progress
    def _right(self, parent) -> None:
        c = self._card(parent, "Progress", expand=True)
        stepper = tk.Frame(c, bg=CARD)
        stepper.pack(fill="x")
        self.step_w = {}
        for i, (key, label) in enumerate(STEPS):
            cell = tk.Frame(stepper, bg=CARD)
            cell.pack(side="left", expand=True, fill="x")
            dot = tk.Label(cell, text=str(i + 1), bg=TINT, fg=MUTED, font=self.F["h"], width=3)
            dot.pack()
            txt = tk.Label(cell, text=label, bg=CARD, fg=MUTED, font=self.F["small"])
            txt.pack(pady=(3, 0))
            self.step_w[key] = (dot, txt, i + 1)
        self.status = tk.Label(c, text="Ready. Fill in steps 1–3, then click Download reports.", bg=CARD, fg=TEXT,
                               font=self.F["body"], anchor="w", justify="left", wraplength=400)
        self.status.pack(fill="x", pady=(12, 6))
        self.bar = ttk.Progressbar(c, style="PF.Horizontal.TProgressbar", maximum=100)
        self.bar.pack(fill="x")

        tbl = tk.Frame(c, bg=CARD)
        tbl.pack(fill="both", expand=True, pady=(12, 8))
        self.tree = ttk.Treeview(tbl, columns=("status",), show="tree headings", style="PF.Treeview", height=5)
        self.tree.heading("#0", text="REPORT", anchor="w")
        self.tree.heading("status", text="STATUS", anchor="w")
        self.tree.column("#0", width=240, stretch=True)
        self.tree.column("status", width=170, anchor="w")
        for tag, col in (("ok", OK), ("bad", BAD), ("busy", PRIMARY), ("idle", MUTED), ("warn", WARN)):
            self.tree.tag_configure(tag, foreground=col)
        self.tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(tbl, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")

        tk.Label(c, text="Activity", bg=CARD, fg=MUTED, font=self.F["small"], anchor="w").pack(fill="x")
        self.logbox = tk.Text(c, height=5, bg=TINT, fg=MUTED, font=self.F["mono"], relief="flat", wrap="word",
                              padx=8, pady=6, state="disabled")
        self.logbox.pack(fill="both")

    # ------------------------------------------------------------ small helpers
    def _client_picked(self, _e=None) -> None:
        for r in self.settings.get("recent_clients") or []:
            if r["client_name"] == self.client.get():
                self.admin.set(r["admin_user"])
                break

    def _select_all(self, on: bool) -> None:
        for v in self.report_vars.values():
            v.set(on)

    def _update_count(self) -> None:
        n = sum(v.get() for v in self.report_vars.values())
        self.count_lbl.config(text=f"{n} of {len(self.report_vars)} selected")

    def _browse(self) -> None:
        d = filedialog.askdirectory(initialdir=self.folder.get() or str(Path.home()), title="Save reports in…")
        if d:
            self.folder.delete(0, "end")
            self.folder.insert(0, os.path.normpath(d))

    def _open_folder(self) -> None:
        path = Path(self.folder.get().strip() or r"C:\Reports download")
        client = self.client.get().strip()
        if client:
            try:
                self.engine.set_download_root(path)
                cand = self.engine.client_download_dir(client)
                if cand.exists():
                    path = cand
            except Exception:
                pass
        path.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{path}"')
        except Exception:
            pass

    def _open_logs(self) -> None:
        """Opens the folder with the run logs and debug page snapshots."""
        path = cfgmod.DATA_DIR
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{path}"')
        except Exception:
            pass

    def _log(self, msg: str) -> None:
        self.q.put(("log", msg))
        try:
            self.log_file.parent.mkdir(parents=True, exist_ok=True)
            with open(self.log_file, "a", encoding="utf-8") as f:
                f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n")
        except Exception:
            pass

    def _append_log(self, msg: str) -> None:
        self.logbox.config(state="normal")
        self.logbox.insert("end", f"{datetime.now():%H:%M:%S}  {msg}\n")
        lines = int(self.logbox.index("end-1c").split(".")[0])
        if lines > 400:
            self.logbox.delete("1.0", f"{lines - 400}.0")
        self.logbox.see("end")
        self.logbox.config(state="disabled")

    def _refresh_signin(self) -> None:
        email = self.engine.CONFIG.get("user_email")
        if self.engine.has_valid_session():
            self.signin_lbl.config(text=f"Signed in{f' ({email})' if email else ''}", fg=OK)
            self.login_btn.config(text="Sign in again")
        else:
            self.signin_lbl.config(text="Not signed in yet. A browser opens for your Google sign-in.", fg=MUTED)

    def _set_step(self, active: str) -> None:
        keys = [k for k, _ in STEPS]
        idx = keys.index(active) if active in keys else (len(keys) if active == "done" else -1)
        for i, k in enumerate(keys):
            dot, txt, num = self.step_w[k]
            if idx == -1 or i > idx:
                dot.config(text=str(num), bg=TINT, fg=MUTED)
                txt.config(fg=MUTED)
            elif i < idx:
                dot.config(text="✓", bg=OK, fg=ON_PRIMARY)
                txt.config(fg=TEXT)
            else:
                dot.config(text=str(num), bg=PRIMARY, fg=ON_PRIMARY)
                txt.config(fg=PRIMARY)

    def _row(self, name: str, status: str, extra: str = "") -> None:
        label, tag = {
            "Waiting": ("Waiting", "idle"),
            "Generating": ("Generating…", "busy"),
            "Downloading": ("Downloading…", "busy"),
            "Success": ("Downloaded", "ok"),
            "Failed": ("Failed", "bad"),
            "Stopped": ("Stopped", "warn"),
        }.get(status, (status, "idle"))
        if extra and status == "Failed":
            label = f"Failed: {extra[:38]}"
        if self.tree.exists(name):
            self.tree.item(name, values=(label,), tags=(tag,))
        else:
            self.tree.insert("", "end", iid=name, text=name, values=(label,), tags=(tag,))

    def _progress_update(self) -> None:
        rows = self.tree.get_children()
        done = sum(1 for r in rows if self.tree.item(r, "values")[0].startswith(("Downloaded", "Failed", "Stopped")))
        if rows:
            self.bar.stop()
            self.bar.config(mode="determinate", value=100 * done / len(rows))

    def _lock_ui(self, busy: str | None) -> None:
        self.busy = busy
        self._set_enabled(self.run_btn, busy is None)
        self._set_enabled(self.login_btn, busy is None)
        self._set_enabled(self.stop_btn, busy == "run")
        self.run_btn.config(text="Working…" if busy == "run" else "Download reports")

    def _save_settings(self, client: str, admin: str, reports: list[str]) -> None:
        s = cfgmod.load_config()
        recent = [r for r in s.get("recent_clients") or [] if r.get("client_name", "").lower() != client.lower()]
        recent.insert(0, {"client_name": client, "admin_user": admin})
        s["recent_clients"] = recent[:15]
        s[self.sel_key] = reports
        s["download_root"] = self.folder.get().strip()
        s["show_browser"] = bool(self.show_browser.get())
        if self.role is not None:
            s["rms_role"] = self.role.get().strip() or "RMG Lead"
        s["notify_on"] = bool(self.notify_on.get())
        s["notify_to"] = self.notify_to.get().strip()
        try:
            cfgmod.save_config(s)
        except Exception:
            pass
        self.settings = s
        self.client.config(values=[r["client_name"] for r in recent[:15]])
        self.admin.config(values=sorted({r["admin_user"] for r in recent[:15]}))

    # ------------------------------------------------------------ actions
    def do_login(self) -> None:
        if self.busy:
            return
        self._lock_ui("login")
        self.status.config(text="A browser window is opening. Complete the Google sign-in there "
                                "(it waits up to 5 minutes).")
        self.signin_lbl.config(text="Signing in…", fg=PRIMARY)

        def work():
            ok = False
            try:
                ok = self.engine.login(log=self._log)
            except Exception as exc:
                self._log(f"ERROR: {exc}")
            self.q.put(("login_done", ok))

        threading.Thread(target=work, daemon=True).start()

    def do_run(self) -> None:
        if self.busy:
            return
        client, admin = self.client.get().strip(), self.admin.get().strip()
        reports = [n for n, v in self.report_vars.items() if v.get()]
        folder = self.folder.get().strip()
        if not self.engine.has_valid_session():
            messagebox.showinfo("Sign in first", "Click 'Sign in' (Step 1) once before downloading reports.")
            return
        if not client or not admin:
            messagebox.showwarning("Client & admin user", "Fill in both Client Name and Admin User (Step 2).")
            return
        if not reports:
            messagebox.showwarning("Reports", "Tick at least one report (Step 3).")
            return
        if not folder:
            messagebox.showwarning("Save to", "Choose a folder to save the reports in.")
            return
        try:
            send_email, recips = self._email_choice()
        except ValueError:
            return
        if send_email:
            self.engine.CONFIG["notify_emails"] = recips
            self.engine.CONFIG["alert_to_email"] = recips[0]
        self._save_settings(client, admin, reports)
        self.engine.set_download_root(folder)
        try:
            created = not Path(folder).is_dir()
            self.engine.ensure_download_root()
            if created:
                self._append_log(f"Created the save folder: {folder}")
        except self.engine.FolderError as exc:
            messagebox.showerror("Save to", str(exc))
            return
        if os.environ.get("PF_POST_TRIGGER_WAIT_S"):
            self.engine.set_post_trigger_wait(int(os.environ["PF_POST_TRIGGER_WAIT_S"]))

        self.tree.delete(*self.tree.get_children())
        for n in reports:
            self._row(n, "Waiting")
        self.result.config(text="")
        self.bar.config(mode="indeterminate")
        self.bar.start(12)
        self._set_step("login")
        self.status.config(text=f"Opening the browser and logging in as {admin} for {client}…")
        self._lock_ui("run")
        self.stop_event = threading.Event()
        headless = not self.show_browser.get()
        extra = {"role": self.role.get().strip() or "RMG Lead"} if self.role is not None else {}

        def work():
            try:
                res = self.runner(client, admin, reports, log=self._log, headless=headless,
                                  stop_event=self.stop_event, trigger="Manual", send_email=send_email, **extra)
                self.q.put(("run_done", ("ok", res)))
            except self.engine.RunStopped:
                self.q.put(("run_done", ("stopped", None)))
            except self.engine.SessionExpired:
                self.q.put(("run_done", ("expired", None)))
            except Exception as exc:
                self.q.put(("run_done", ("error", str(exc))))

        threading.Thread(target=work, daemon=True).start()

    def do_stop(self) -> None:
        if self.stop_event:
            self.stop_event.set()
            self.status.config(text="Stopping after the current step… (can take a few seconds)")
            self._set_enabled(self.stop_btn, False)

    # ------------------------------------------------------------ event loop
    def _poll(self) -> None:
        try:
            while True:
                kind, data = self.q.get_nowait()
                self._handle(kind, data)
        except queue.Empty:
            pass
        if self.wait_since:
            el = int(time.time() - self.wait_since)
            self.status.config(text=f"Waiting for '{self.wait_label}' to finish on the server… "
                                    f"{el // 60}:{el % 60:02d} so far (checks every 15s)")
        if self.wait_until:
            left = int(self.wait_until - time.time())
            if left > 0:
                self.status.config(text=f"All reports triggered. Waiting for the server to build them… "
                                        f"{left // 60}:{left % 60:02d} left")
            else:
                self.wait_until = 0
        self.root.after(200, self._poll)

    def _handle(self, kind: str, data) -> None:
        if kind == "log":
            msg = data
            self._append_log(msg)
            if msg.startswith("Logged in as"):
                self.status.config(text="Logged in. Opening the reports page…")
            elif msg.startswith("PHASE 1"):
                self._set_step("generate")
                self.status.config(text="Starting report generation…")
            elif msg.startswith("PHASE 2"):
                self._set_step("wait")
                self.wait_until = time.time() + int(self.engine.POST_TRIGGER_WAIT_S)
            elif msg.startswith("REPORT "):  # IntelliPrism: "REPORT 2/5 Positions Report"
                head, _, name = msg[7:].partition(" ")
                self._set_step("generate")
                self.wait_since = 0
                self.status.config(text=f"Report {head.replace('/', ' of ')}: {name}. Setting dates and generating…")
            elif msg.startswith("WAITING "):
                self._set_step("wait")
                self.wait_since = time.time()
                self.wait_label = msg[8:]
            elif msg.startswith("[") and "Saved to" in msg and self.product == "intelliprism":
                self._set_step("download")
            elif msg.startswith("PHASE 3"):
                self._set_step("download")
                self.wait_until = 0
                self.status.config(text="Downloading fresh files…")
                for r in self.tree.get_children():
                    if self.tree.item(r, "values")[0].startswith("Generating"):
                        self._row(r, "Downloading")
                self._progress_update()
        elif kind == "status":
            report, status, file_path, error = data
            if self.product == "intelliprism" and status in ("Success", "Failed", "Stopped"):
                self.wait_since = 0
            self._row(report, status, error or "")
            self._progress_update()
        elif kind == "login_done":
            self._lock_ui(None)
            self._refresh_signin()
            self.status.config(text="Signed in. Now fill in Step 2 and 3." if data else
                               "Sign-in wasn't completed. Click Sign in to try again.")
        elif kind == "run_done":
            self._finish(*data)

    def _finish(self, outcome: str, payload) -> None:
        self._lock_ui(None)
        self.wait_until = 0
        self.wait_since = 0
        self.bar.stop()
        self.bar.config(mode="determinate", value=100)
        rows = self.tree.get_children()
        ok = sum(1 for r in rows if self.tree.item(r, "values")[0] == "Downloaded")
        folder = self.folder.get().strip()
        if outcome == "ok":
            self._set_step("done")
            if ok == len(rows):
                self.result.config(text=f"✓  All done: {ok} report(s) downloaded.", fg=OK)
            elif ok:
                self.result.config(text=f"!  Partly done: {ok} of {len(rows)} downloaded. See the failed rows.",
                                   fg=WARN)
            else:
                self.result.config(text="✕  Nothing was downloaded. See the failed rows and Activity.", fg=BAD)
            self.status.config(text=f"Saved under {folder}\\<Client>\\<date>  ·  log: Download_Log.xlsx")
            if ok:
                self._open_folder()
        elif outcome == "stopped":
            self._set_step("")
            self.result.config(text="Run stopped.", fg=WARN)
        elif outcome == "expired":
            self._set_step("")
            self.result.config(text="✕  Your sign-in has expired. Click 'Sign in again', then retry.", fg=BAD)
            self._refresh_signin()
        else:
            self._set_step("")
            self.result.config(text=f"✕  {str(payload)[:160]}", fg=BAD)
        if os.environ.get("PF_SELFTEST"):
            self.root.after(300, self.root.destroy)

    def _on_close(self) -> None:
        if self.busy and not messagebox.askyesno("Still working", "A run is in progress. Close anyway?"):
            return
        if self.stop_event:
            self.stop_event.set()
        self.root.destroy()

    def _selftest(self) -> None:
        """Proves the packaged app works: UI built, engine + Playwright driver present."""
        from playwright._impl._driver import compute_driver_executable

        drv = compute_driver_executable()
        drv_path = Path(drv[0] if isinstance(drv, tuple) else drv)
        report = cfgmod.DATA_DIR / f"selftest_{self.product}.txt"
        sched = "skipped (not Windows)"
        if self.sched.supported():
            try:  # create a real Windows task far in the future, check it, then remove it
                t = self.sched.add_schedule(self.product, "SelfTest", "selftest@example.com", ["x"],
                                            str(cfgmod.DATA_DIR / "selftest_out"), "03:17", ["SUN"])
                made = self.sched.task_exists(t)
                self.sched.remove_schedule(t["id"])
                gone = not self.sched.task_exists(t)
                sched = "ok" if made and gone else f"FAILED made={made} gone={gone}"
            except Exception as exc:
                sched = f"FAILED {exc}"
        kr = "skipped (not Windows)"
        if self.sched.supported():
            try:  # the Gmail app password must land in Windows Credential Manager
                import keyring

                keyring.set_password("PF Report Downloader selftest", "x", "secret-123")
                got = keyring.get_password("PF Report Downloader selftest", "x")
                keyring.delete_password("PF Report Downloader selftest", "x")
                backend = type(keyring.get_keyring()).__name__
                kr = "ok" if got == "secret-123" and ("WinVault" in backend or "Windows" in backend) else f"FAILED backend={backend} got={bool(got)}"
            except Exception as exc:
                kr = f"FAILED {exc}"
        report.write_text(f"ui=ok\nproduct={self.product}\nreports={len(self.report_vars)}\ndriver={drv_path}\n"
                          f"driver_exists={drv_path.exists()}\nschedule={sched}\nkeyring={kr}\n", encoding="utf-8")
        self.root.destroy()


def main(product: str = "skillprism") -> int:
    if sys.stdout is None or sys.stderr is None:  # windowed .exe: no console
        logf = open(cfgmod.DATA_DIR / "app_output.txt", "a", encoding="utf-8", buffering=1)  # noqa: SIM115
        sys.stdout = sys.stdout or logf
        sys.stderr = sys.stderr or logf
    load_brand_font()
    root = tk.Tk()
    App(root, product)
    root.mainloop()
    return 0


def run_safely(product: str) -> None:
    """Run the app; a windowed .exe must never fail silently."""
    try:
        sys.exit(main(product))
    except SystemExit:
        raise
    except Exception as exc:
        import traceback

        details = traceback.format_exc()
        try:
            (cfgmod.DATA_DIR / "error_log.txt").write_text(details, encoding="utf-8")
        except Exception:
            pass
        try:
            messagebox.showerror("Prismforce Report Downloader",
                                 f"The app hit an error:\n\n{exc}\n\nDetails: {cfgmod.DATA_DIR / 'error_log.txt'}")
        except Exception:
            pass
        sys.exit(1)


if __name__ == "__main__":
    run_safely(os.environ.get("PF_PRODUCT", "skillprism"))
