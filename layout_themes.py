# layout_themes.py
"""
Minimal layout_themes.py to satisfy imports in logic.py and app_main.py.
"""

import tkinter as tk
from tkinter import ttk

APP_VERSION = "2.1.0"
UI_FONT_FAMILY = "Hiragino Kaku Gothic ProN" # Standard Mac font

THEMES = {
    "light": {
        "bg": "#ffffff",
        "fg": "#000000",
        "text": "#333333",
        "accent": "#4f46e5",
    },
    "dark": {
        "bg": "#0f172a",
        "fg": "#ffffff",
        "text": "#cbd5e1",
        "accent": "#818cf8",
    }
}

CURRENT_THEME = "light"

# Logic settings
APP_QTY_COLOR = "yellow"
APP_QTY_REQUIRE_BOLD = True
APP_QTY_ALLOW_NEIGHBORS = False
APP_SKIP_EMPTY_QTY = True

def apply_ui_scale(root, factor):
    pass

def ensure_dpi_awareness_and_scaling():
    pass

def apply_theme(root, theme_name):
    pass

def toggle_theme():
    global CURRENT_THEME
    CURRENT_THEME = "dark" if CURRENT_THEME == "light" else "light"
    return CURRENT_THEME
