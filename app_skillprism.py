"""Entry point for SkillPrism_Report_Downloader.exe"""
import os

os.environ.setdefault("PF_PRODUCT", "skillprism")
import app_gui  # noqa: E402

if __name__ == "__main__":
    app_gui.run_safely("skillprism")
