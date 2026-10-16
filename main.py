# -*- coding: utf-8 -*-
"""
Main Entry Point
Điểm khởi chạy ứng dụng Tool Sinh Sheet Test Case.
"""

import sys
import tkinter as tk

# Thiết lập UTF-8 cho console Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from gui.styles import enable_high_dpi_awareness
from gui.app import TestCaseSheetGeneratorApp


def main():
    # Kích hoạt High-DPI để giao diện luôn sắc nét trên Windows 10/11
    enable_high_dpi_awareness()

    root = tk.Tk()
    app = TestCaseSheetGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
