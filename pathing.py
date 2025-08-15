# utils/pathing.py
import sys
from pathlib import Path

def resource_path(*parts: str) -> str:
    """
    Returns a path that works both in dev and in PyInstaller.
    Use: resource_path("templates") or resource_path("templates", "pay_report_template.html")
    """
    base = getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1])  # adjust if utils/ is nested differently
    return str(Path(base, *parts))
