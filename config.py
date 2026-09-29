import os
from pathlib import Path
import customtkinter as ctk

# Application Appearance
ctk.set_appearance_mode("dark")

# Project Root and Local Data Directory
PROJECT_ROOT = Path(__file__).parent
DATA_DIR = PROJECT_ROOT / "data"

ENTITY_A_AC_DIR = DATA_DIR / "ENTITY_A AC"
ENTITY_B_AC_DIR = DATA_DIR / "ENTITY_B AC"
MASTER_AC_DIR = DATA_DIR / "MASTER AC"

# Ensure all local data directories exist automatically on startup
for folder in [ENTITY_A_AC_DIR, ENTITY_B_AC_DIR, MASTER_AC_DIR]:
    os.makedirs(folder, exist_ok=True)

# Cosmic Space Tech Theme Palette
BG_MAIN = "#131A2E"
BG_FRAME = "#1E2942"
BORDER_COLOR = "#00F0FF"

ACCENT_CYAN = "#00F0FF"
ACCENT_PURPLE = "#365EE2"
ACCENT_PURPLE_H = "#274ABD"

COLOR_SUCCESS = "#00E676"
COLOR_SUCCESS_H = "#00B454"
COLOR_DANGER = "#FF2D55"
COLOR_DANGER_H = "#D31B40"

TEXT_WHITE = "#E2E8F0"
TEXT_MUTED = "#94A3B8"