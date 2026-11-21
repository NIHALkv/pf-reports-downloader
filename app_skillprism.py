"""Entry point for SkillPrism_Report_Downloader.exe

    (double-click)          -> opens the app window
    --scheduled <id>        -> background run of a saved schedule (started by Windows Task Scheduler)
"""
import os
import sys

os.environ.setdefault("PF_PRODUCT", "skillprism")

if __name__ == "__main__":
    if "--scheduled" in sys.argv:
        import scheduling

        sched_id = sys.argv[sys.argv.index("--scheduled") + 1]
        sys.exit(scheduling.run_scheduled("skillprism", sched_id))
    import app_gui

    app_gui.run_safely("skillprism")
