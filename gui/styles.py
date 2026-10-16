# -*- coding: utf-8 -*-
"""
GUI Styles & Theme Configuration
Thiết lập theme hiện đại, bảng màu Slate/Navy, font chữ Segoe UI và hỗ trợ màn hình High-DPI.
"""

import sys
import tkinter as tk
from tkinter import ttk

# ================= COLOR PALETTE =================
BG_MAIN = "#F1F5F9"        # Nền chính (Slate-100)
BG_CARD = "#FFFFFF"        # Nền Card / Khung (Trắng)
BG_SUBTLE = "#F8FAFC"      # Nền phụ (Slate-50)
BORDER_COLOR = "#CBD5E1"   # Viền (Slate-300)
BORDER_LIGHT = "#E2E8F0"   # Viền mờ (Slate-200)

TEXT_MAIN = "#0F172A"      # Chữ chính (Slate-900)
TEXT_MUTED = "#64748B"     # Chữ phụ (Slate-500)
TEXT_LIGHT = "#94A3B8"     # Chữ mờ (Slate-400)

COLOR_PRIMARY = "#1E40AF"  # Xanh Navy (Blue-800)
COLOR_PRIMARY_HOVER = "#1D4ED8"
COLOR_ACCENT = "#2563EB"   # Xanh hoàng gia (Blue-600)
COLOR_SUCCESS = "#059669"  # Xanh lục (Emerald-600)
COLOR_SUCCESS_HOVER = "#047857"
COLOR_WARNING = "#D97706"  # Vàng cam (Amber-600)
COLOR_DANGER = "#DC2626"   # Đỏ (Red-600)

# ================= FONTS =================
FONT_FAMILY = "Segoe UI"
FONT_TITLE = (FONT_FAMILY, 15, "bold")
FONT_SUBTITLE = (FONT_FAMILY, 9)
FONT_SECTION = (FONT_FAMILY, 11, "bold")
FONT_BODY = (FONT_FAMILY, 9)
FONT_BODY_BOLD = (FONT_FAMILY, 9, "bold")
FONT_SMALL = (FONT_FAMILY, 8)
FONT_CODE = ("Consolas", 9)


def enable_high_dpi_awareness():
    """Bật chế độ DPI Awareness trên Windows để chữ và icon sắc nét."""
    if sys.platform == "win32":
        try:
            import ctypes
            # Shcore DPI Awareness
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass


def setup_theme(root: tk.Tk):
    """Thiết lập ttk styles hiện đại cho ứng dụng."""
    style = ttk.Style(root)
    available_themes = style.theme_names()
    if "clam" in available_themes:
        style.theme_use("clam")

    # Base Frame
    style.configure("TFrame", background=BG_MAIN)
    style.configure("Card.TFrame", background=BG_CARD, relief="flat")
    style.configure("Header.TFrame", background=COLOR_PRIMARY)

    # Base Labels
    style.configure("TLabel", background=BG_MAIN, foreground=TEXT_MAIN, font=FONT_BODY)
    style.configure("Card.TLabel", background=BG_CARD, foreground=TEXT_MAIN, font=FONT_BODY)
    style.configure("CardMuted.TLabel", background=BG_CARD, foreground=TEXT_MUTED, font=FONT_BODY)
    style.configure("CardTitle.TLabel", background=BG_CARD, foreground=COLOR_PRIMARY, font=FONT_SECTION)
    style.configure("CardDesc.TLabel", background=BG_CARD, foreground=TEXT_MUTED, font=FONT_SUBTITLE)
    style.configure("HeaderTitle.TLabel", background=COLOR_PRIMARY, foreground="#FFFFFF", font=FONT_TITLE)
    style.configure("HeaderSub.TLabel", background=COLOR_PRIMARY, foreground="#BFDBFE", font=FONT_SUBTITLE)
    style.configure("Badge.TLabel", background="#E0E7FF", foreground=COLOR_PRIMARY, font=FONT_BODY_BOLD, padding=4)

    # Buttons
    style.configure(
        "Primary.TButton",
        background=COLOR_PRIMARY,
        foreground="#FFFFFF",
        font=FONT_BODY_BOLD,
        borderwidth=0,
        focuscolor="none",
        padding=(14, 8)
    )
    style.map(
        "Primary.TButton",
        background=[("active", COLOR_PRIMARY_HOVER), ("disabled", "#94A3B8")],
        foreground=[("disabled", "#F1F5F9")]
    )

    style.configure(
        "Success.TButton",
        background=COLOR_SUCCESS,
        foreground="#FFFFFF",
        font=FONT_BODY_BOLD,
        borderwidth=0,
        focuscolor="none",
        padding=(16, 9)
    )
    style.map(
        "Success.TButton",
        background=[("active", COLOR_SUCCESS_HOVER), ("disabled", "#94A3B8")]
    )

    style.configure(
        "Secondary.TButton",
        background="#E2E8F0",
        foreground=TEXT_MAIN,
        font=FONT_BODY,
        borderwidth=0,
        focuscolor="none",
        padding=(10, 6)
    )
    style.map(
        "Secondary.TButton",
        background=[("active", "#CBD5E1")]
    )

    style.configure(
        "Action.TButton",
        background="#EFF6FF",
        foreground=COLOR_PRIMARY,
        font=FONT_BODY_BOLD,
        borderwidth=1,
        bordercolor=COLOR_PRIMARY,
        focuscolor="none",
        padding=(10, 5)
    )
    style.map(
        "Action.TButton",
        background=[("active", "#DBEAFE")]
    )

    # Checkbutton & Radiobutton
    style.configure(
        "TCheckbutton",
        background=BG_CARD,
        foreground=TEXT_MAIN,
        font=FONT_BODY,
        focuscolor="none"
    )
    style.map("TCheckbutton", background=[("active", BG_CARD)])

    style.configure(
        "TRadiobutton",
        background=BG_CARD,
        foreground=TEXT_MAIN,
        font=FONT_BODY,
        focuscolor="none"
    )
    style.map("TRadiobutton", background=[("active", BG_CARD)])

    # Combobox
    style.configure(
        "TCombobox",
        font=FONT_BODY,
        padding=4
    )

    # Treeview
    style.configure(
        "Treeview",
        background="#FFFFFF",
        foreground=TEXT_MAIN,
        fieldbackground="#FFFFFF",
        rowheight=26,
        font=FONT_BODY
    )
    style.configure(
        "Treeview.Heading",
        background="#F1F5F9",
        foreground=TEXT_MAIN,
        font=FONT_BODY_BOLD,
        padding=5
    )
    style.map(
        "Treeview.Heading",
        background=[("active", "#E2E8F0")]
    )
    style.map(
        "Treeview",
        background=[("selected", "#DBEAFE")],
        foreground=[("selected", COLOR_PRIMARY)]
    )

    # Progressbar
    style.configure(
        "Horizontal.TProgressbar",
        background=COLOR_ACCENT,
        troughcolor="#E2E8F0",
        borderwidth=0,
        thickness=8
    )

    # Labelframe
    style.configure(
        "TLabelframe",
        background=BG_CARD,
        bordercolor=BORDER_LIGHT,
        borderwidth=1,
        relief="solid"
    )
    style.configure(
        "TLabelframe.Label",
        background=BG_CARD,
        foreground=COLOR_PRIMARY,
        font=FONT_SECTION
    )

    # Notebook (Tabs)
    style.configure(
        "TNotebook",
        background=BG_CARD,
        borderwidth=0
    )
    style.configure(
        "TNotebook.Tab",
        background="#E2E8F0",
        foreground=TEXT_MAIN,
        font=FONT_BODY_BOLD,
        padding=(16, 8)
    )
    style.map(
        "TNotebook.Tab",
        background=[("selected", BG_CARD), ("active", "#CBD5E1")],
        foreground=[("selected", COLOR_PRIMARY)]
    )
