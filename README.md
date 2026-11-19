# SkillPrism Report Downloader

Download Reports Center reports for any client in a few clicks: sign in, pick the client and admin user, tick the reports, and click **Download reports**.

![SkillPrism Report Downloader](docs/app-screenshot.png)

**[⬇️ Download SkillPrism_Report_Downloader.exe](../../releases/latest/download/SkillPrism_Report_Downloader.exe)** (Windows, nothing to install)

---

## How to use
1. **Open** `SkillPrism_Report_Downloader.exe`.
   - Chrome says *"isn't commonly downloaded"*: click **⋯ → Keep → Keep anyway**.
   - Windows says *"protected your PC"*: click **More info → Run anyway**. The app is unsigned.
2. **Step 1: Sign in** (first time only). An Edge window opens and clicks *Sign in with Google* for you. Pick your @prismforce.ai account and finish 2FA if asked. The sign-in is remembered.
3. **Step 2: Client & admin user.** Type them exactly as in **Global Admin → Clients** and that client's **Users** list. The admin user needs **Silent Login** access. Recent clients are remembered in the drop-downs.
4. **Step 3: Reports.** Tick reports by category, or use **Select all**.
5. Click **Download reports**. Progress shows live: **Silent login → Generate → Wait → Download**, with each report's status. **Stop** cancels between steps.

Files are saved to `C:\Reports download\<Client>\<DD-MM-YYYY>\<Report>_<HHMMSS>.xlsx`. Change the folder with **Save to**. Every attempt is logged in `Download_Log.xlsx` in that folder.

## Good to know
- **Nothing to set up:** no Setup.bat, no config.json, no command prompt. Your email is picked up from the sign-in.
- **Passwords are never entered or stored by the app.** Google sign-in happens in the browser window.
- **Show the browser while it works:** untick to run in the background.
- **Settings, sign-in and logs** live in `%APPDATA%\PF Reports Downloader\`:
  - `run_logs\`: a log of every run
  - `debug\`: a page snapshot whenever a step fails. Send it along when reporting a problem.
  - `error_log.txt`: written if the app itself crashes
- **Optional email summaries:** add `"gmail_app_password": "<16 chars>"` to `%APPDATA%\PF Reports Downloader\config.json`. Without it, emails are skipped.
- Only one run at a time. Close any other copy of the app first, because they share the signed-in browser.

## How it works
`automation.py` is the proven engine from the SkillPrism Reports Downloader (desktop edition):

Global Admin → search client → **Silent Login** as the admin user → Reports Center → **Phase 1:** trigger *Generate Report* for every selected report → **Phase 2:** wait 2 min → **Phase 3:** download only the Download History rows created by this run.

| File | What it is |
|---|---|
| `app_gui.py` | The branded desktop app. This is what the .exe runs |
| `automation.py` | Playwright engine: sign-in, silent login, trigger, download |
| `config.py` | Settings (auto-created in `%APPDATA%` for the .exe) |
| `report_log.py` | Excel run log (Summary / By Day / Log sheets) |
| `branding/` | Official Prismforce logo and icon from the brand kit |
| `.github/workflows/build.yml` | Builds the .exe on Windows, self-tests it, publishes the release |

### Run from source
```bash
pip install -r requirements.txt
python app_gui.py
```

### Releasing a new version
Push to `main`. GitHub builds `SkillPrism_Report_Downloader.exe` on Windows, launches it once as a self-test (window opens, engine and browser driver load), and only then publishes it to **[Releases → latest](../../releases/latest)**. The download link above always serves the newest build.
