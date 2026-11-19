"""Entry point for IntelliPrism_Report_Downloader.exe"""
import os

os.environ.setdefault("PF_PRODUCT", "intelliprism")
import app_gui  # noqa: E402

if __name__ == "__main__":
    app_gui.run_safely("intelliprism")
