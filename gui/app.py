# -*- coding: utf-8 -*-
"""
Main Application Window
Giao diện Desktop hiện đại xây dựng bằng Tkinter và ttk.
"""

import os
import sys
import threading
import subprocess
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from typing import Dict, List, Any, Optional

from core.excel_processor import (
    inspect_excel_file,
    get_sheet_columns_preview,
    analyze_groups,
    generate_excel_sheets,
    sanitize_sheet_name,
    download_google_sheet_as_excel
)
from core.sample_generator import create_sample_master_file
from gui.styles import (
    enable_high_dpi_awareness,
    setup_theme,
    BG_MAIN,
    BG_CARD,
    BG_SUBTLE,
    BORDER_COLOR,
    BORDER_LIGHT,
    TEXT_MAIN,
    TEXT_MUTED,
    COLOR_PRIMARY,
    COLOR_ACCENT,
    COLOR_SUCCESS,
    COLOR_DANGER,
    FONT_FAMILY,
    FONT_TITLE,
    FONT_SUBTITLE,
    FONT_SECTION,
    FONT_BODY,
    FONT_BODY_BOLD,
    FONT_CODE,
    FONT_SMALL
)


class TestCaseSheetGeneratorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Công Cụ Sinh Sheet Test Case Tự Động - v2.0 Enterprise QA")
        self.root.geometry("1340x850")
        self.root.minsize(1080, 700)
        self.root.configure(bg=BG_MAIN)

        setup_theme(self.root)

        # Trạng thái dữ liệu
        self.current_file_path: str = ""
        self.available_sheets: List[str] = []
        self.columns_info: List[Dict[str, Any]] = []
        self.preview_raw_rows: List[List[str]] = []
        self.detected_groups: Dict[str, Dict[str, Any]] = {}
        self.selected_group_names: set = set()
        self.is_processing: bool = False

        # Build UI
        self._create_header()
        self._create_body()

        # Tự động nạp file mẫu nếu có sẵn
        self._check_default_sample()

    def _create_header(self):
        """Tạo thanh tiêu đề phía trên."""
        header_frame = tk.Frame(self.root, bg=COLOR_PRIMARY, height=75)
        header_frame.pack(fill="x", side="top")
        header_frame.pack_propagate(False)

        inner_frame = tk.Frame(header_frame, bg=COLOR_PRIMARY)
        inner_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # Tiêu đề bên trái
        title_box = tk.Frame(inner_frame, bg=COLOR_PRIMARY)
        title_box.pack(side="left", fill="y")

        lbl_title = tk.Label(
            title_box,
            text="📑 TOOL SINH SHEET TEST CASE TỰ ĐỘNG",
            font=FONT_TITLE,
            bg=COLOR_PRIMARY,
            fg="#FFFFFF"
        )
        lbl_title.pack(anchor="w")

        lbl_sub = tk.Label(
            title_box,
            text="Tự động bóc tách từng test case con thành từng Sheet độc lập | Giữ nguyên 100% Template, Header đa tầng & Màu sắc",
            font=FONT_SUBTITLE,
            bg=COLOR_PRIMARY,
            fg="#BFDBFE"
        )
        lbl_sub.pack(anchor="w", pady=(2, 0))

        # Nút tiện ích bên phải
        btn_box = tk.Frame(inner_frame, bg=COLOR_PRIMARY)
        btn_box.pack(side="right", fill="y")

        btn_sample = tk.Button(
            btn_box,
            text="⚡ Nạp File Mẫu Chuẩn",
            font=FONT_BODY_BOLD,
            bg="#2563EB",
            fg="#FFFFFF",
            activebackground="#1D4ED8",
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            padx=12,
            pady=6,
            command=self._on_generate_sample_clicked
        )
        btn_sample.pack(side="right", pady=5)

        btn_webapp = tk.Button(
            btn_box,
            text="🌐 Mở Bản Web App (Trình Duyệt)",
            font=FONT_BODY_BOLD,
            bg="#0D9488",
            fg="#FFFFFF",
            activebackground="#0F766E",
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            padx=12,
            pady=6,
            command=self._on_open_web_app
        )
        btn_webapp.pack(side="right", padx=(0, 10), pady=5)

    def _create_body(self):
        """Tạo layout nội dung chính gồm 2 cột: Cấu hình (trái) và Không gian dữ liệu (phải)."""
        main_container = tk.Frame(self.root, bg=BG_MAIN)
        main_container.pack(fill="both", expand=True, padx=16, pady=14)

        # Cột trái 390px, cột phải chiếm phần lớn không gian màn hình
        main_container.columnconfigure(0, weight=3, minsize=410)
        main_container.columnconfigure(1, weight=7, minsize=680)
        main_container.rowconfigure(0, weight=1)

        # Cột Trái (Cấu hình 3 bước)
        left_col = tk.Frame(main_container, bg=BG_MAIN)
        left_col.grid(row=0, column=0, sticky="nsew", padx=(0, 12))

        # Cột Phải (Bảng xem trước & Điều khiển)
        right_col = tk.Frame(main_container, bg=BG_MAIN)
        right_col.grid(row=0, column=1, sticky="nsew")

        self._create_step1_card(left_col)
        self._create_step2_card(left_col)
        self._create_step3_card(left_col)

        self._create_preview_card(right_col)
        self._create_action_card(right_col)

    def _create_step1_card(self, parent):
        """Card Bước 1: Chọn File và Sheet nguồn."""
        card = self._make_card(parent, "📁 BƯỚC 1: NHẬP FILE EXCEL HOẶC GOOGLE SHEETS")
        card.pack(fill="x", pady=(0, 10))

        # Chọn hình thức nguồn (File Excel / Google Sheet)
        src_mode_frame = tk.Frame(card, bg=BG_CARD)
        src_mode_frame.pack(fill="x", padx=12, pady=(6, 4))

        self.var_source_type = tk.StringVar(value="file")

        rb_file = ttk.Radiobutton(
            src_mode_frame,
            text="📁 File Excel từ máy (.xlsx, .xls)",
            variable=self.var_source_type,
            value="file",
            command=self._on_source_type_changed
        )
        rb_file.pack(side="left", padx=(0, 14))

        rb_gsheet = ttk.Radiobutton(
            src_mode_frame,
            text="🌐 Link Google Sheets",
            variable=self.var_source_type,
            value="gsheet",
            command=self._on_source_type_changed
        )
        rb_gsheet.pack(side="left")

        # Frame chứa input file máy
        self.frame_file_input = tk.Frame(card, bg=BG_CARD)
        self.frame_file_input.pack(fill="x", padx=12, pady=(2, 6))

        tk.Label(self.frame_file_input, text="Đường dẫn file Excel:", font=FONT_BODY_BOLD, bg=BG_CARD, fg=TEXT_MAIN).pack(anchor="w")
        f_select = tk.Frame(self.frame_file_input, bg=BG_CARD)
        f_select.pack(fill="x", pady=(3, 0))

        self.txt_file_path = ttk.Entry(f_select, font=FONT_BODY)
        self.txt_file_path.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.txt_file_path.bind("<Return>", lambda e: self._on_manual_file_path_entered())

        btn_browse = ttk.Button(f_select, text="Chọn File...", style="Primary.TButton", command=self._on_browse_file)
        btn_browse.pack(side="right")

        # Frame chứa input Google Sheet URL
        self.frame_gsheet_input = tk.Frame(card, bg=BG_CARD)

        tk.Label(self.frame_gsheet_input, text="Link Google Sheets (Yêu cầu bật 'Bất kỳ ai có liên kết'):", font=FONT_BODY_BOLD, bg=BG_CARD, fg=TEXT_MAIN).pack(anchor="w")
        g_select = tk.Frame(self.frame_gsheet_input, bg=BG_CARD)
        g_select.pack(fill="x", pady=(3, 0))

        self.txt_gsheet_url = ttk.Entry(g_select, font=FONT_BODY)
        self.txt_gsheet_url.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.txt_gsheet_url.insert(0, "https://docs.google.com/spreadsheets/d/...")
        self.txt_gsheet_url.bind("<Return>", lambda e: self._on_download_gsheet())

        btn_gsheet = ttk.Button(g_select, text="📥 Tải & Đọc", style="Primary.TButton", command=self._on_download_gsheet)
        btn_gsheet.pack(side="right")

        # Dòng chọn/nhập sheet và dòng header
        s_row = tk.Frame(card, bg=BG_CARD)
        s_row.pack(fill="x", padx=12, pady=(0, 10))

        # Chọn hoặc Nhập Tên Sheet
        s_left = tk.Frame(s_row, bg=BG_CARD)
        s_left.pack(side="left", fill="x", expand=True, padx=(0, 8))

        tk.Label(s_left, text="Tên Sheet chứa Test Case (Chọn hoặc Gõ tên):", font=FONT_BODY, bg=BG_CARD, fg=TEXT_MAIN).pack(anchor="w")
        self.cbo_sheet = ttk.Combobox(s_left, state="normal", font=FONT_BODY)
        self.cbo_sheet.pack(fill="x", pady=(2, 0))
        self.cbo_sheet.bind("<<ComboboxSelected>>", self._on_sheet_selected)
        self.cbo_sheet.bind("<Return>", lambda e: self._on_sheet_selected(None))

        # Chọn dòng header
        s_right = tk.Frame(s_row, bg=BG_CARD)
        s_right.pack(side="right", fill="x")

        tk.Label(s_right, text="Dòng tiêu đề (Header):", font=FONT_BODY, bg=BG_CARD, fg=TEXT_MAIN).pack(anchor="w")
        self.spin_header_row = ttk.Spinbox(s_right, from_=1, to=10, width=8, font=FONT_BODY)
        self.spin_header_row.set("1")
        self.spin_header_row.pack(pady=(2, 0))
        self.spin_header_row.bind("<FocusOut>", lambda e: self._on_header_row_changed())
        self.spin_header_row.bind("<Return>", lambda e: self._on_header_row_changed())

    def _create_step2_card(self, parent):
        """Card Bước 2: Cấu hình quy tắc tách sheet."""
        card = self._make_card(parent, "⚙️ BƯỚC 2: QUY TẮC TÁCH SHEET TỪNG TEST CASE")
        card.pack(fill="x", pady=(0, 10))

        inner = tk.Frame(card, bg=BG_CARD)
        inner.pack(fill="x", padx=12, pady=(8, 10))

        # Chọn quy tắc tách
        tk.Label(inner, text="Quy tắc sinh Sheet:", font=FONT_BODY_BOLD, bg=BG_CARD, fg=TEXT_MAIN).pack(anchor="w")

        rule_frame = tk.Frame(inner, bg=BG_CARD)
        rule_frame.pack(fill="x", pady=(2, 8))

        self.var_split_rule = tk.StringVar(value="each_case")

        rb_case = ttk.Radiobutton(
            rule_frame,
            text="🎯 Tách từng Case riêng lẻ (Mỗi case = 1 Sheet)",
            variable=self.var_split_rule,
            value="each_case",
            command=self._on_split_rule_changed
        )
        rb_case.pack(anchor="w", pady=(0, 2))

        rb_group = ttk.Radiobutton(
            rule_frame,
            text="📁 Gom theo Module / Phân nhóm",
            variable=self.var_split_rule,
            value="by_group",
            command=self._on_split_rule_changed
        )
        rb_group.pack(anchor="w")

        # Chọn cột
        self.lbl_col_group = tk.Label(inner, text="Cột định danh Test Case (dùng làm Tên Sheet):", font=FONT_BODY_BOLD, bg=BG_CARD, fg=TEXT_MAIN)
        self.lbl_col_group.pack(anchor="w", pady=(6, 2))

        c_box = tk.Frame(inner, bg=BG_CARD)
        c_box.pack(fill="x", pady=(0, 8))

        self.cbo_group_col = ttk.Combobox(c_box, state="readonly", font=FONT_BODY)
        self.cbo_group_col.pack(side="left", fill="x", expand=True, padx=(0, 8))

        btn_analyze = ttk.Button(c_box, text="🔍 Quét Case", style="Action.TButton", command=self._analyze_file_groups)
        btn_analyze.pack(side="right")

        # Tùy chọn Forward fill cho merged cells
        self.var_forward_fill = tk.BooleanVar(value=True)
        chk_ff = ttk.Checkbutton(
            inner,
            text="Tự động xử lý ô gộp (Merged Cells / Điền tiếp cho dòng con nhiều bước)",
            variable=self.var_forward_fill,
            style="TCheckbutton"
        )
        chk_ff.pack(anchor="w", pady=(0, 4))

        self.lbl_rule_hint = tk.Label(
            inner,
            text="💡 Mỗi test case sẽ sinh 1 sheet độc lập với template giống hệt sheet tổng.",
            font=FONT_SMALL,
            bg=BG_CARD,
            fg=TEXT_MUTED
        )
        self.lbl_rule_hint.pack(anchor="w")

    def _create_step3_card(self, parent):
        """Card Bước 3: Tùy chọn xuất file."""
        card = self._make_card(parent, "🛠️ BƯỚC 3: TÙY CHỌN XUẤT SHEET & ĐỊNH DẠNG")
        card.pack(fill="x", pady=(0, 0))

        inner = tk.Frame(card, bg=BG_CARD)
        inner.pack(fill="x", padx=12, pady=(8, 10))

        # Checkbox Tổng quan
        self.var_summary = tk.BooleanVar(value=True)
        chk_summary = ttk.Checkbutton(
            inner,
            text="Tạo Sheet 'Tổng quan' (Overview) kèm bảng điều hướng nhấp mở nhanh",
            variable=self.var_summary,
            style="TCheckbutton"
        )
        chk_summary.pack(anchor="w", pady=(0, 4))

        # Checkbox Back link
        self.var_back_link = tk.BooleanVar(value=True)
        chk_back = ttk.Checkbutton(
            inner,
            text="Thêm liên kết quay lại trang Tổng quan ở dòng đầu mỗi sheet con",
            variable=self.var_back_link,
            style="TCheckbutton"
        )
        chk_back.pack(anchor="w", pady=(0, 4))

        # Checkbox Giữ Styles
        self.var_styles = tk.BooleanVar(value=True)
        chk_styles = ttk.Checkbutton(
            inner,
            text="Giữ nguyên định dạng màu sắc ô, font chữ, viền và độ rộng cột từ file gốc",
            variable=self.var_styles,
            style="TCheckbutton"
        )
        chk_styles.pack(anchor="w", pady=(0, 8))

        # Chế độ xuất (Single file vs Multiple files)
        tk.Label(inner, text="Hình thức xuất:", font=FONT_BODY_BOLD, bg=BG_CARD, fg=TEXT_MAIN).pack(anchor="w", pady=(2, 2))

        self.var_export_mode = tk.StringVar(value="single_workbook")
        rb_frame = tk.Frame(inner, bg=BG_CARD)
        rb_frame.pack(fill="x", pady=(0, 8))

        rb_single = ttk.Radiobutton(
            rb_frame,
            text="Gộp chung vào 1 file Excel (Nhiều Sheet)",
            variable=self.var_export_mode,
            value="single_workbook",
            command=self._on_export_mode_changed
        )
        rb_single.pack(side="left", padx=(0, 16))

        rb_multi = ttk.Radiobutton(
            rb_frame,
            text="Tách mỗi nhóm thành 1 file riêng lẻ (.xlsx)",
            variable=self.var_export_mode,
            value="multiple_files",
            command=self._on_export_mode_changed
        )
        rb_multi.pack(side="left")

        # Đường dẫn file/thư mục xuất
        self.lbl_output_target = tk.Label(inner, text="File Excel kết quả:", font=FONT_BODY_BOLD, bg=BG_CARD, fg=TEXT_MAIN)
        self.lbl_output_target.pack(anchor="w", pady=(4, 2))

        out_box = tk.Frame(inner, bg=BG_CARD)
        out_box.pack(fill="x")

        self.txt_output_path = ttk.Entry(out_box, font=FONT_BODY)
        self.txt_output_path.pack(side="left", fill="x", expand=True, padx=(0, 8))

        btn_out_browse = ttk.Button(out_box, text="Chọn nơi lưu...", style="Secondary.TButton", command=self._on_browse_output)
        btn_out_browse.pack(side="right")

    def _create_preview_card(self, parent):
        """Card Khung làm việc dữ liệu đa tab: Bảng dữ liệu chi tiết 20 cột, Danh sách sheet sẽ tách, và Xem trước sheet con."""
        card = self._make_card(parent, "📊 KHÔNG GIAN DỮ LIỆU & XEM TRƯỚC SHEET")
        card.pack(fill="both", expand=True, pady=(0, 10))

        # Khởi tạo Notebook
        self.notebook = ttk.Notebook(card)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=8)

        # Tab 1: Bảng dữ liệu chi tiết
        self.tab_data = tk.Frame(self.notebook, bg=BG_CARD)
        self.notebook.add(self.tab_data, text="  📊 Bảng Dữ Liệu Test Case Chi Tiết  ")

        # Tab 2: Danh sách sheet sẽ tạo
        self.tab_sheets = tk.Frame(self.notebook, bg=BG_CARD)
        self.notebook.add(self.tab_sheets, text="  📑 Danh Sách Sheet Sẽ Tách (Mỗi Case = 1 Sheet)  ")

        # Tab 3: Xem trước Sheet con
        self.tab_single_preview = tk.Frame(self.notebook, bg=BG_CARD)
        self.notebook.add(self.tab_single_preview, text="  👁️ Xem Trước Mẫu Sheet Con  ")

        self._build_tab_data(self.tab_data)
        self._build_tab_sheets(self.tab_sheets)
        self._build_tab_single_preview(self.tab_single_preview)

    def _build_tab_data(self, parent):
        """Tab 1: Bảng dữ liệu test case đầy đủ 20 cột chuẩn QA."""
        top_bar = tk.Frame(parent, bg=BG_CARD)
        top_bar.pack(fill="x", padx=10, pady=(8, 6))

        self.lbl_data_stats = tk.Label(
            top_bar,
            text="Chưa nạp dữ liệu",
            font=FONT_BODY_BOLD,
            bg="#EFF6FF",
            fg=COLOR_PRIMARY,
            padx=8,
            pady=3
        )
        self.lbl_data_stats.pack(side="left")

        # Khung tìm kiếm nhanh
        search_box = tk.Frame(top_bar, bg=BG_CARD)
        search_box.pack(side="right")

        tk.Label(search_box, text="🔍 Lọc:", font=FONT_BODY, bg=BG_CARD, fg=TEXT_MUTED).pack(side="left", padx=(0, 4))
        self.txt_data_search = ttk.Entry(search_box, width=22, font=FONT_BODY)
        self.txt_data_search.pack(side="left", padx=(0, 4))
        self.txt_data_search.bind("<KeyRelease>", self._on_search_data)

        btn_clear_search = ttk.Button(search_box, text="✕", width=3, style="Secondary.TButton", command=self._clear_search_data)
        btn_clear_search.pack(side="left")

        # Bảng hiển thị dữ liệu có 2 thanh cuộn
        grid_frame = tk.Frame(parent, bg=BG_CARD)
        grid_frame.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        self.tree_data = ttk.Treeview(grid_frame, show="headings", selectmode="browse")

        v_scroll = ttk.Scrollbar(grid_frame, orient="vertical", command=self.tree_data.yview)
        h_scroll = ttk.Scrollbar(grid_frame, orient="horizontal", command=self.tree_data.xview)

        self.tree_data.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        self.tree_data.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")

        grid_frame.rowconfigure(0, weight=1)
        grid_frame.columnconfigure(0, weight=1)

        self.tree_data.tag_configure("even", background="#FFFFFF")
        self.tree_data.tag_configure("odd", background="#F8FAFC")

        self.tree_data.bind("<Double-Button-1>", self._on_data_row_double_click)

        # Chú thích nhỏ dưới bảng
        lbl_hint = tk.Label(
            parent,
            text="💡 Nhấp đúp chuột vào bất kỳ dòng nào để xem chi tiết toàn bộ các bước và dữ liệu của test case.",
            font=FONT_SMALL,
            bg=BG_CARD,
            fg=TEXT_MUTED
        )
        lbl_hint.pack(anchor="w", padx=10, pady=(0, 6))

    def _build_tab_sheets(self, parent):
        """Tab 2: Danh sách các sheet con dự kiến sinh ra."""
        tb = tk.Frame(parent, bg=BG_CARD)
        tb.pack(fill="x", padx=10, pady=(8, 6))

        self.lbl_group_summary = tk.Label(
            tb,
            text="Chưa quét dữ liệu",
            font=FONT_BODY_BOLD,
            bg="#EFF6FF",
            fg=COLOR_PRIMARY,
            padx=8,
            pady=3
        )
        self.lbl_group_summary.pack(side="left")

        # Nút chọn tất cả / bỏ chọn / đổi tên
        btn_rename = ttk.Button(tb, text="✏️ Đổi tên Sheet", style="Secondary.TButton", command=self._on_rename_sheet)
        btn_rename.pack(side="right", padx=(4, 0))

        btn_select_all = ttk.Button(tb, text="✓ Chọn tất cả", style="Secondary.TButton", command=self._select_all_groups)
        btn_select_all.pack(side="right", padx=(4, 0))

        btn_deselect_all = ttk.Button(tb, text="✗ Bỏ chọn", style="Secondary.TButton", command=self._deselect_all_groups)
        btn_deselect_all.pack(side="right")

        # Bảng Treeview danh sách sheet
        tree_frame = tk.Frame(parent, bg=BG_CARD)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        columns = ("selected", "stt", "group_name", "sheet_name", "count", "status")
        self.tree_groups = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")

        self.tree_groups.heading("selected", text="Chọn")
        self.tree_groups.heading("stt", text="STT")
        self.tree_groups.heading("group_name", text="Mã Test Case / Nhóm Gốc")
        self.tree_groups.heading("sheet_name", text="Tên Sheet Sẽ Tạo")
        self.tree_groups.heading("count", text="Số TC Con")
        self.tree_groups.heading("status", text="Trạng Thái")

        self.tree_groups.column("selected", width=55, anchor="center")
        self.tree_groups.column("stt", width=45, anchor="center")
        self.tree_groups.column("group_name", width=220, anchor="w")
        self.tree_groups.column("sheet_name", width=200, anchor="w")
        self.tree_groups.column("count", width=80, anchor="center")
        self.tree_groups.column("status", width=110, anchor="center")

        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree_groups.yview)
        self.tree_groups.configure(yscrollcommand=v_scroll.set)

        self.tree_groups.pack(side="left", fill="both", expand=True)
        v_scroll.pack(side="right", fill="y")

        self.tree_groups.bind("<ButtonRelease-1>", self._on_tree_click)
        self.tree_groups.bind("<Double-Button-1>", self._on_tree_double_click)

        lbl_tbl_hint = tk.Label(
            parent,
            text="* Nhấp vào dòng để Bật/Tắt chọn Sheet. Nhấn 'Đổi tên Sheet' nếu muốn sửa tên sheet theo ý muốn.",
            font=FONT_SMALL,
            bg=BG_CARD,
            fg=TEXT_MUTED
        )
        lbl_tbl_hint.pack(anchor="w", padx=10, pady=(0, 6))

    def _build_tab_single_preview(self, parent):
        """Tab 3: Mô phỏng giao diện 1 Sheet con khi được sinh ra."""
        top_ctrl = tk.Frame(parent, bg=BG_CARD)
        top_ctrl.pack(fill="x", padx=10, pady=(8, 6))

        tk.Label(top_ctrl, text="Chọn Test Case muốn xem trước mẫu Sheet:", font=FONT_BODY_BOLD, bg=BG_CARD, fg=TEXT_MAIN).pack(side="left", padx=(0, 8))

        self.cbo_preview_case = ttk.Combobox(top_ctrl, state="readonly", width=26, font=FONT_BODY)
        self.cbo_preview_case.pack(side="left", padx=(0, 8))
        self.cbo_preview_case.bind("<<ComboboxSelected>>", lambda e: self._render_single_sheet_preview())

        btn_refresh_preview = ttk.Button(top_ctrl, text="🔄 Xem mẫu", style="Action.TButton", command=self._render_single_sheet_preview)
        btn_refresh_preview.pack(side="left")

        # Khung mô phỏng bảng tính
        frame_mockup = tk.Frame(parent, bg=BORDER_LIGHT, padx=1, pady=1)
        frame_mockup.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        inner_mockup = tk.Frame(frame_mockup, bg="#FFFFFF")
        inner_mockup.pack(fill="both", expand=True)

        # Tab bar giả lập trong Excel
        mock_tab_bar = tk.Frame(inner_mockup, bg="#F1F5F9", height=32)
        mock_tab_bar.pack(fill="x", side="top")
        self.lbl_mock_sheet_title = tk.Label(
            mock_tab_bar,
            text="📄 Sheet: [Chọn Test Case để xem trước]",
            font=FONT_BODY_BOLD,
            bg="#FFFFFF",
            fg=COLOR_PRIMARY,
            padx=12,
            pady=4,
            relief="solid",
            borderwidth=1
        )
        self.lbl_mock_sheet_title.pack(side="left", padx=8, pady=(4, 0))

        # Vùng hiển thị cấu trúc Sheet
        f_txt = tk.Frame(inner_mockup, bg="#FFFFFF")
        f_txt.pack(fill="both", expand=True, padx=8, pady=8)

        self.txt_single_preview = tk.Text(f_txt, font=FONT_BODY, wrap="none", padx=12, pady=10, bg="#FFFFFF", relief="flat")
        v_sp = ttk.Scrollbar(f_txt, orient="vertical", command=self.txt_single_preview.yview)
        h_sp = ttk.Scrollbar(f_txt, orient="horizontal", command=self.txt_single_preview.xview)
        self.txt_single_preview.configure(yscrollcommand=v_sp.set, xscrollcommand=h_sp.set)

        self.txt_single_preview.grid(row=0, column=0, sticky="nsew")
        v_sp.grid(row=0, column=1, sticky="ns")
        h_sp.grid(row=1, column=0, sticky="ew")

        f_txt.rowconfigure(0, weight=1)
        f_txt.columnconfigure(0, weight=1)

        self.txt_single_preview.tag_configure("header_cyan", font=FONT_BODY_BOLD, background="#00B0F0", foreground="#000000")
        self.txt_single_preview.tag_configure("banner_yellow", font=FONT_BODY_BOLD, background="#FFE699", foreground="#000000")
        self.txt_single_preview.tag_configure("case_data", font=FONT_BODY, background="#FFFFFF", foreground="#000000")
        self.txt_single_preview.tag_configure("bold", font=FONT_BODY_BOLD)

        lbl_mock_note = tk.Label(
            parent,
            text="✅ Mỗi Sheet con khi sinh ra sẽ giữ nguyên vẹn 100% Header xanh Cyan (#00B0F0), Banner vàng (#FFE699) và chỉ bóc tách dữ liệu của riêng Case này.",
            font=FONT_SMALL,
            bg=BG_CARD,
            fg=COLOR_SUCCESS
        )
        lbl_mock_note.pack(anchor="w", padx=10, pady=(0, 6))

    def _render_data_table(self, rows=None):
        """Hiển thị các dòng dữ liệu lên Treeview tab 1."""
        if not hasattr(self, "tree_data") or not self.columns_info:
            return

        col_ids = [f"c_{c['index']}" for c in self.columns_info]
        self.tree_data["columns"] = col_ids

        # Cấu hình headings và độ rộng cột
        for c in self.columns_info:
            cid = f"c_{c['index']}"
            name = c["name"]
            self.tree_data.heading(cid, text=f"{c['letter']}: {name}")

            # Đặt độ rộng phù hợp cho từng loại cột
            name_lower = name.lower()
            if "id" in name_lower and "jira" not in name_lower:
                w = 110
            elif "mục đích" in name_lower:
                w = 260
            elif "bước" in name_lower:
                w = 280
            elif "kết quả" in name_lower:
                w = 260
            elif "data test" in name_lower:
                w = 140
            elif "uat" in name_lower or "prod" in name_lower:
                w = 120
            else:
                w = 130
            self.tree_data.column(cid, width=w, minwidth=80, anchor="w")

        self.tree_data.delete(*self.tree_data.get_children())

        display_rows = self.preview_raw_rows if rows is None else rows
        for idx, row in enumerate(display_rows):
            tag = "even" if idx % 2 == 0 else "odd"
            padded = list(row) + [""] * (len(self.columns_info) - len(row))
            self.tree_data.insert("", "end", iid=f"row_{idx}", values=padded[:len(self.columns_info)], tags=(tag,))

        total_cols = len(self.columns_info)
        total_rows = len(display_rows)
        self.lbl_data_stats.config(
            text=f"📊 Hiển thị: {total_rows} dòng | {total_cols} cột chuẩn QA"
        )

    def _on_search_data(self, event=None):
        """Lọc dữ liệu theo từ khóa tìm kiếm."""
        q = self.txt_data_search.get().strip().lower()
        if not q:
            self._render_data_table(self.preview_raw_rows)
            return

        filtered = []
        for r in self.preview_raw_rows:
            row_text = " ".join(str(cell) for cell in r).lower()
            if q in row_text:
                filtered.append(r)

        self._render_data_table(filtered)

    def _clear_search_data(self):
        """Xóa ô tìm kiếm và nạp lại toàn bộ."""
        self.txt_data_search.delete(0, tk.END)
        self._render_data_table(self.preview_raw_rows)

    def _on_data_row_double_click(self, event):
        """Mở cửa sổ xem chi tiết dòng testcase được nhấp đúp."""
        item_id = self.tree_data.focus()
        if not item_id:
            return
        vals = self.tree_data.item(item_id, "values")
        if not vals:
            return

        dlg = tk.Toplevel(self.root)
        tc_name = vals[0] if len(vals) > 0 and vals[0] else "Dòng dữ liệu"
        dlg.title(f"Chi Tiết Test Case - {tc_name}")
        dlg.geometry("780x620")
        dlg.minsize(620, 480)
        dlg.configure(bg=BG_CARD)

        header_dlg = tk.Frame(dlg, bg=COLOR_PRIMARY, height=45)
        header_dlg.pack(fill="x")
        tk.Label(header_dlg, text=f"📋 THÔNG TIN CHI TIẾT: {tc_name}", font=FONT_SECTION, bg=COLOR_PRIMARY, fg="#FFFFFF").pack(side="left", padx=16, pady=10)

        f_content = tk.Frame(dlg, bg=BG_CARD)
        f_content.pack(fill="both", expand=True, padx=16, pady=12)

        txt_info = tk.Text(f_content, font=FONT_BODY, wrap="word", padx=12, pady=12, bg="#F8FAFC", relief="solid", borderwidth=1)
        v_s = ttk.Scrollbar(f_content, orient="vertical", command=txt_info.yview)
        txt_info.configure(yscrollcommand=v_s.set)
        txt_info.pack(side="left", fill="both", expand=True)
        v_s.pack(side="right", fill="y")

        txt_info.tag_configure("bold_title", font=FONT_BODY_BOLD, foreground=COLOR_PRIMARY)
        txt_info.tag_configure("content", font=FONT_BODY, foreground=TEXT_MAIN)

        for idx, col in enumerate(self.columns_info):
            c_name = col["name"]
            c_val = vals[idx] if idx < len(vals) else ""
            txt_info.insert("end", f"▶ {col['letter']}. {c_name}:\n", "bold_title")
            txt_info.insert("end", f"{c_val if c_val else '(Trống)'}\n\n", "content")

        txt_info.config(state="disabled")

        btn_close = ttk.Button(dlg, text="Đóng", style="Primary.TButton", command=dlg.destroy)
        btn_close.pack(pady=(0, 12))

    def _update_single_preview(self):
        """Cập nhật danh sách test case vào combobox của Tab 3."""
        if hasattr(self, "cbo_preview_case") and self.detected_groups:
            cases = list(self.detected_groups.keys())
            self.cbo_preview_case["values"] = cases
            if cases:
                self.cbo_preview_case.current(0)
                self._render_single_sheet_preview()

    def _render_single_sheet_preview(self):
        """Vẽ lại mô phỏng của 1 Sheet con được chọn."""
        if not hasattr(self, "txt_single_preview"):
            return

        selected_case = self.cbo_preview_case.get()
        if not selected_case and self.detected_groups:
            selected_case = list(self.detected_groups.keys())[0]

        if not selected_case or selected_case not in self.detected_groups:
            return

        group_info = self.detected_groups[selected_case]
        target_sheet_name = group_info["sheet_name"]
        self.lbl_mock_sheet_title.config(text=f"📄 Sheet: {target_sheet_name} (Chuẩn bị xuất)")

        self.txt_single_preview.config(state="normal")
        self.txt_single_preview.delete("1.0", tk.END)

        self.txt_single_preview.insert("end", f"=== MÔ PHỎNG NỘI DUNG SHEET CON: '{target_sheet_name}' ===\n\n", "bold")
        self.txt_single_preview.insert("end", "⬅ Quay lại Trang Tổng quan (Overview)\n\n", "bold")

        # 1. Header 20 cột (Cyan #00B0F0)
        self.txt_single_preview.insert("end", "[DÒNG TIÊU ĐỀ 1 & 2 - MÀU XANH CYAN #00B0F0 (MERGED HEADERS)]\n", "header_cyan")
        hdr_line = " | ".join(f"{c['letter']}: {c['name']}" for c in self.columns_info)
        self.txt_single_preview.insert("end", f"{hdr_line}\n\n", "header_cyan")

        # 2. Banner vàng
        self.txt_single_preview.insert("end", "[DÒNG PHÂN HỆ BANNER - MÀU VÀNG #FFE699]\n", "banner_yellow")
        self.txt_single_preview.insert("end", "Job quét và tổng hợp danh sách nợ (Merged toàn bộ các cột từ A đến T)\n\n", "banner_yellow")

        # 3. Dòng test case con
        self.txt_single_preview.insert("end", f"[DỮ LIỆU TEST CASE CON: {selected_case}]\n", "bold")

        # Tìm dòng trong preview_raw_rows
        found_rows = [r for r in self.preview_raw_rows if len(r) > 0 and r[0] == selected_case]
        if not found_rows and self.preview_raw_rows:
            found_rows = [self.preview_raw_rows[0]]

        for r_idx, row in enumerate(found_rows, start=1):
            self.txt_single_preview.insert("end", f"--- Dòng bước {r_idx} ---\n", "bold")
            for c_idx, col in enumerate(self.columns_info):
                c_val = row[c_idx] if c_idx < len(row) else ""
                if c_val:
                    self.txt_single_preview.insert("end", f"  • {col['name']}: ", "bold")
                    self.txt_single_preview.insert("end", f"{c_val}\n", "case_data")
            self.txt_single_preview.insert("end", "\n")

        self.txt_single_preview.config(state="disabled")

    def _on_rename_sheet(self):
        """Cho phép người dùng tùy chỉnh tên sheet sẽ tạo."""
        selected_item = self.tree_groups.focus()
        if not selected_item:
            messagebox.showinfo("Chọn nhóm", "Vui lòng nhấp chọn 1 dòng trong danh sách bên dưới để đổi tên Sheet.")
            return

        current_info = self.detected_groups.get(selected_item)
        if not current_info:
            return

        old_sheet_name = current_info["sheet_name"]
        new_name = simpledialog.askstring(
            "Đổi tên Sheet",
            f"Nhập tên Sheet mới cho Test Case '{selected_item}':",
            initialvalue=old_sheet_name,
            parent=self.root
        )
        if new_name and new_name.strip():
            clean_name = sanitize_sheet_name(new_name.strip(), [])
            self.detected_groups[selected_item]["sheet_name"] = clean_name
            current_vals = list(self.tree_groups.item(selected_item, "values"))
            current_vals[3] = clean_name
            self.tree_groups.item(selected_item, values=current_vals)
            self._update_single_preview()

    def _on_open_web_app(self):
        """Mở giao diện Web App trên trình duyệt."""
        web_file = os.path.abspath("web/index.html")
        if os.path.exists(web_file):
            webbrowser.open(f"file://{web_file}")
        else:
            messagebox.showinfo("Web App", "Không tìm thấy file web/index.html")

    def _create_action_card(self, parent):
        """Card Tiến trình và Nút bấm sinh sheet."""
        card = self._make_card(parent, "🚀 THỰC THI & TIẾN TRÌNH")
        card.pack(fill="x", pady=(0, 0))

        inner = tk.Frame(card, bg=BG_CARD)
        inner.pack(fill="x", padx=12, pady=(8, 12))

        # Progress bar và Status label
        prog_info = tk.Frame(inner, bg=BG_CARD)
        prog_info.pack(fill="x", pady=(0, 4))

        self.lbl_status = tk.Label(prog_info, text="Sẵn sàng thực hiện", font=FONT_BODY, bg=BG_CARD, fg=TEXT_MUTED)
        self.lbl_status.pack(side="left")

        self.lbl_percent = tk.Label(prog_info, text="0%", font=FONT_BODY_BOLD, bg=BG_CARD, fg=COLOR_PRIMARY)
        self.lbl_percent.pack(side="right")

        self.progress_bar = ttk.Progressbar(inner, orient="horizontal", mode="determinate", style="Horizontal.TProgressbar")
        self.progress_bar.pack(fill="x", pady=(0, 10))

        # Nút bấm hành động chính
        self.btn_run = tk.Button(
            inner,
            text="🚀 BẮT ĐẦU SINH SHEET TEST CASE",
            font=(FONT_FAMILY, 12, "bold"),
            bg=COLOR_SUCCESS,
            fg="#FFFFFF",
            activebackground=COLOR_SUCCESS,
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            padx=20,
            pady=10,
            command=self._on_start_generate
        )
        self.btn_run.pack(fill="x")

    def _make_card(self, parent, title: str) -> tk.Frame:
        """Helper tạo khung card bo viền đẹp mắt."""
        outer = tk.Frame(parent, bg=BORDER_LIGHT, padx=1, pady=1)
        card = tk.Frame(outer, bg=BG_CARD)
        card.pack(fill="both", expand=True)

        # Title bar của card
        t_frame = tk.Frame(card, bg=BG_SUBTLE, height=32)
        t_frame.pack(fill="x", side="top")
        t_frame.pack_propagate(False)

        lbl = tk.Label(t_frame, text=title, font=FONT_SECTION, bg=BG_SUBTLE, fg=COLOR_PRIMARY)
        lbl.pack(side="left", padx=12, pady=6)

        div = tk.Frame(card, bg=BORDER_LIGHT, height=1)
        div.pack(fill="x", side="top")

        return card

    # ================= LOGIC & SỰ KIỆN =================

    def _check_default_sample(self):
        """Nếu có sẵn file mẫu trong thư mục sample_data, tự động nạp để người dùng dễ thử nghiệm."""
        default_sample = os.path.abspath("sample_data/sample_master_testcases.xlsx")
        if os.path.exists(default_sample):
            self.txt_file_path.delete(0, tk.END)
            self.txt_file_path.insert(0, default_sample)
            self.spin_header_row.set("3")
            self._load_file_metadata(default_sample)

    def _on_generate_sample_clicked(self):
        """Tạo file test case mẫu và tự động nạp vào tool."""
        try:
            target_path = os.path.abspath("sample_data/sample_master_testcases.xlsx")
            created_path = create_sample_master_file(target_path)
            self.txt_file_path.delete(0, tk.END)
            self.txt_file_path.insert(0, created_path)
            self.spin_header_row.set("3")
            self._load_file_metadata(created_path)

            ans = messagebox.askyesno(
                "Đã nạp file mẫu thành công!",
                f"Đã nạp file testcase mẫu 20 cột chuẩn ISO/IEC IEEE 29119:\n{created_path}\n\nBạn có muốn mở file này lên xem bằng Excel không?"
            )
            if ans:
                self._open_file_or_dir(created_path)
        except Exception as e:
            messagebox.showerror("Lỗi tạo file mẫu", f"Không thể tạo file mẫu:\n{str(e)}")

    def _on_source_type_changed(self):
        """Chuyển đổi giữa chế độ File máy và Link Google Sheet."""
        stype = self.var_source_type.get()
        if stype == "file":
            self.frame_gsheet_input.pack_forget()
            self.frame_file_input.pack(fill="x", padx=12, pady=(2, 6))
        else:
            self.frame_file_input.pack_forget()
            self.frame_gsheet_input.pack(fill="x", padx=12, pady=(2, 6))

    def _on_manual_file_path_entered(self):
        """Khi người dùng gõ hoặc dán đường dẫn file và nhấn Enter."""
        path = self.txt_file_path.get().strip().strip('"').strip("'")
        if os.path.exists(path):
            self._load_file_metadata(path)
        else:
            messagebox.showwarning("Không tìm thấy file", f"Đường dẫn file không tồn tại:\n{path}")

    def _on_download_gsheet(self):
        """Tải file từ Google Sheet và nạp vào tool."""
        url = self.txt_gsheet_url.get().strip()
        if not url or "docs.google.com" not in url:
            messagebox.showwarning("Link không hợp lệ", "Vui lòng nhập link Google Sheet hợp lệ dạng:\nhttps://docs.google.com/spreadsheets/d/.../edit")
            return

        self.lbl_status.config(text="Đang kết nối và tải Google Sheet...", fg=COLOR_PRIMARY)
        self.root.update_idletasks()

        target_save = os.path.abspath("sample_data/google_sheet_download.xlsx")
        try:
            download_google_sheet_as_excel(url, target_save)
            self.txt_file_path.delete(0, tk.END)
            self.txt_file_path.insert(0, target_save)
            self._load_file_metadata(target_save)
            self.lbl_status.config(text="Đã tải và nạp Google Sheet thành công!", fg=COLOR_SUCCESS)
            messagebox.showinfo("Thành công", f"Đã tải thành công dữ liệu từ Google Sheet!\nBạn có thể chọn cột phân nhóm và bấm Sinh Sheet ngay.")
        except Exception as e:
            self.lbl_status.config(text="Lỗi tải Google Sheet", fg=COLOR_DANGER)
            messagebox.showerror("Lỗi tải Google Sheet", str(e))

    def _on_browse_file(self):
        """Mở hộp thoại chọn file Excel nguồn."""
        filetypes = [
            ("File Excel (*.xlsx;*.xls)", "*.xlsx;*.xls"),
            ("Tất cả các file (*.*)", "*.*")
        ]
        selected = filedialog.askopenfilename(
            title="Chọn file Test Case tổng",
            filetypes=filetypes
        )
        if selected:
            self.txt_file_path.delete(0, tk.END)
            self.txt_file_path.insert(0, selected)
            self._load_file_metadata(selected)

    def _load_file_metadata(self, file_path: str):
        """Đọc danh sách sheet và thiết lập các thông tin ban đầu."""
        try:
            if not os.path.exists(file_path):
                return
            self.current_file_path = file_path
            info = inspect_excel_file(file_path)
            self.available_sheets = info["sheet_names"]

            self.cbo_sheet["values"] = self.available_sheets
            if self.available_sheets:
                self.cbo_sheet.current(0)
                self._on_sheet_selected(None)

            # Đề xuất đường dẫn file xuất
            self._suggest_output_path(file_path)
        except Exception as e:
            messagebox.showerror("Lỗi đọc file", f"Không thể đọc file Excel:\n{str(e)}")

    def _suggest_output_path(self, input_path: str):
        """Tự động sinh tên file xuất gợi ý."""
        dir_name = os.path.dirname(input_path)
        base_name, ext = os.path.splitext(os.path.basename(input_path))
        mode = self.var_export_mode.get()

        if mode == "single_workbook":
            suggested = os.path.join(dir_name, f"{base_name}_Tach_Sheet{ext}")
            self.txt_output_path.delete(0, tk.END)
            self.txt_output_path.insert(0, suggested)
        else:
            suggested_dir = os.path.join(dir_name, f"{base_name}_Sheets_Xuat")
            self.txt_output_path.delete(0, tk.END)
            self.txt_output_path.insert(0, suggested_dir)

    def _on_sheet_selected(self, event):
        """Khi người dùng chọn sheet khác."""
        self._load_columns_for_current_sheet()

    def _on_header_row_changed(self):
        """Khi người dùng thay đổi dòng Header."""
        self._load_columns_for_current_sheet()

    def _on_split_rule_changed(self):
        """Khi chuyển đổi giữa Tách từng Case và Gom theo Module."""
        rule = self.var_split_rule.get()
        if rule == "each_case":
            self.lbl_col_group.config(text="Cột định danh Test Case (dùng làm Tên Sheet):")
            self.lbl_rule_hint.config(text="💡 Mỗi test case sẽ sinh 1 sheet độc lập với template giống hệt sheet tổng.")
        else:
            self.lbl_col_group.config(text="Cột dùng để phân nhóm gom testcase:")
            self.lbl_rule_hint.config(text="💡 Gom các testcase cùng nhóm (Module, Feature...) vào chung 1 sheet.")
        self._load_columns_for_current_sheet()

    def _load_columns_for_current_sheet(self):
        """Nạp danh sách các cột từ sheet đang chọn."""
        sheet_name = self.cbo_sheet.get()
        if not self.current_file_path or not sheet_name:
            return

        try:
            h_row = int(self.spin_header_row.get() or "1")
        except ValueError:
            h_row = 1

        try:
            col_data = get_sheet_columns_preview(self.current_file_path, sheet_name, header_row=h_row)
            self.columns_info = col_data["columns"]
            self.preview_raw_rows = col_data.get("preview_rows", [])

            # Cập nhật Bảng dữ liệu chi tiết Tab 1
            self._render_data_table()

            # Cập nhật vào combobox chọn cột
            col_display_list = [f"Cột {c['letter']}: {c['name']}" for c in self.columns_info]
            self.cbo_group_col["values"] = col_display_list

            rule = self.var_split_rule.get()
            best_idx = 0
            if rule == "each_case":
                keywords = ["con", "mã test case", "mã tc", "sub-id", "sub id", "test case id", "mã", "id", "tên test case"]
            else:
                keywords = ["module", "phân hệ", "chức năng", "feature", "cha", "suite", "nhóm", "category"]

            for idx, c in enumerate(self.columns_info):
                c_name_lower = c["name"].lower()
                if any(kw in c_name_lower for kw in keywords):
                    best_idx = idx
                    break

            if col_display_list:
                self.cbo_group_col.current(best_idx)
                # Tự động quét nhóm luôn
                self._analyze_file_groups()

        except Exception as e:
            messagebox.showwarning("Cảnh báo", f"Không thể đọc danh sách cột từ sheet '{sheet_name}':\n{str(e)}")

    def _analyze_file_groups(self):
        """Phân tích các nhóm trong cột đã chọn và hiển thị lên bảng Treeview."""
        file_path = self.txt_file_path.get().strip()
        sheet_name = self.cbo_sheet.get().strip()
        col_selection_idx = self.cbo_group_col.current()

        if not file_path or not os.path.exists(file_path):
            messagebox.showwarning("Thiếu thông tin", "Vui lòng chọn file Excel nguồn hợp lệ.")
            return

        if col_selection_idx < 0 or col_selection_idx >= len(self.columns_info):
            messagebox.showwarning("Thiếu thông tin", "Vui lòng chọn cột phân nhóm.")
            return

        col_obj = self.columns_info[col_selection_idx]
        col_real_idx = col_obj["index"]

        try:
            h_row = int(self.spin_header_row.get() or "1")
        except ValueError:
            h_row = 1

        forward_fill = self.var_forward_fill.get()

        self.lbl_status.config(text="Đang phân tích các nhóm dữ liệu...", fg=COLOR_PRIMARY)
        self.root.update_idletasks()

        try:
            self.detected_groups = analyze_groups(
                file_path=file_path,
                sheet_name=sheet_name,
                group_column_idx=col_real_idx,
                header_row=h_row,
                forward_fill=forward_fill
            )

            # Cập nhật lên Treeview
            self._render_groups_tree()
            self._update_single_preview()

            total_tc = sum(g["count"] for g in self.detected_groups.values())
            self.lbl_group_summary.config(
                text=f"Tìm thấy: {len(self.detected_groups)} test case | Tổng {total_tc} bước"
            )
            self.lbl_status.config(text=f"Đã phân tích xong {len(self.detected_groups)} test case", fg=COLOR_SUCCESS)

        except Exception as e:
            self.lbl_status.config(text="Lỗi khi phân tích nhóm", fg=COLOR_DANGER)
            messagebox.showerror("Lỗi phân tích", f"Không thể phân tích dữ liệu:\n{str(e)}")

    def _render_groups_tree(self):
        """Vẽ lại bảng danh sách nhóm trên Treeview."""
        self.tree_groups.delete(*self.tree_groups.get_children())
        self.selected_group_names = set(self.detected_groups.keys())

        for idx, (group_name, info) in enumerate(self.detected_groups.items(), start=1):
            is_selected = group_name in self.selected_group_names
            mark = "☑" if is_selected else "☐"
            status = "Sẵn sàng tạo"
            self.tree_groups.insert(
                "",
                "end",
                iid=group_name,
                values=(mark, idx, group_name, info["sheet_name"], info["count"], status)
            )

    def _on_tree_click(self, event):
        """Xử lý nhấp chuột vào dòng để bật/tắt chọn."""
        region = self.tree_groups.identify_region(event.x, event.y)
        if region in ("cell", "tree"):
            item_id = self.tree_groups.identify_row(event.y)
            if item_id:
                self._toggle_group_selection(item_id)

    def _on_tree_double_click(self, event):
        """Nhấp đúp chuột cũng bật tắt chọn."""
        item_id = self.tree_groups.focus()
        if item_id:
            self._toggle_group_selection(item_id)

    def _toggle_group_selection(self, group_name: str):
        """Đảo trạng thái chọn của một nhóm."""
        if group_name in self.selected_group_names:
            self.selected_group_names.remove(group_name)
            mark = "☐"
        else:
            self.selected_group_names.add(group_name)
            mark = "☑"

        current_vals = list(self.tree_groups.item(group_name, "values"))
        current_vals[0] = mark
        self.tree_groups.item(group_name, values=current_vals)

        # Cập nhật tóm tắt
        sel_count = len(self.selected_group_names)
        sel_tc = sum(self.detected_groups[g]["count"] for g in self.selected_group_names if g in self.detected_groups)
        self.lbl_group_summary.config(
            text=f"Đã chọn: {sel_count}/{len(self.detected_groups)} nhóm | {sel_tc} test case"
        )

    def _select_all_groups(self):
        """Chọn tất cả các nhóm."""
        self.selected_group_names = set(self.detected_groups.keys())
        for group_name in self.detected_groups.keys():
            if self.tree_groups.exists(group_name):
                current_vals = list(self.tree_groups.item(group_name, "values"))
                current_vals[0] = "☑"
                self.tree_groups.item(group_name, values=current_vals)

        total_tc = sum(g["count"] for g in self.detected_groups.values())
        self.lbl_group_summary.config(
            text=f"Đã chọn: {len(self.detected_groups)}/{len(self.detected_groups)} nhóm | {total_tc} test case"
        )

    def _deselect_all_groups(self):
        """Bỏ chọn tất cả các nhóm."""
        self.selected_group_names.clear()
        for group_name in self.detected_groups.keys():
            if self.tree_groups.exists(group_name):
                current_vals = list(self.tree_groups.item(group_name, "values"))
                current_vals[0] = "☐"
                self.tree_groups.item(group_name, values=current_vals)

        self.lbl_group_summary.config(
            text=f"Đã chọn: 0/{len(self.detected_groups)} nhóm | 0 test case"
        )

    def _on_export_mode_changed(self):
        """Khi người dùng chuyển đổi chế độ xuất."""
        mode = self.var_export_mode.get()
        if mode == "single_workbook":
            self.lbl_output_target.config(text="File Excel kết quả (.xlsx):")
        else:
            self.lbl_output_target.config(text="Thư mục chứa các file Excel xuất ra:")

        if self.current_file_path:
            self._suggest_output_path(self.current_file_path)

    def _on_browse_output(self):
        """Chọn nơi lưu file hoặc thư mục đích."""
        mode = self.var_export_mode.get()
        if mode == "single_workbook":
            selected = filedialog.asksaveasfilename(
                title="Chọn nơi lưu file Excel kết quả",
                defaultextension=".xlsx",
                filetypes=[("File Excel (*.xlsx)", "*.xlsx")]
            )
        else:
            selected = filedialog.askdirectory(title="Chọn thư mục chứa các file Excel xuất ra")

        if selected:
            self.txt_output_path.delete(0, tk.END)
            self.txt_output_path.insert(0, selected)

    def _on_start_generate(self):
        """Bắt đầu tiến trình sinh sheet (chạy luồng phụ để không treo giao diện)."""
        if self.is_processing:
            return

        input_file = self.txt_file_path.get().strip()
        sheet_name = self.cbo_sheet.get().strip()
        output_path = self.txt_output_path.get().strip()
        col_idx_sel = self.cbo_group_col.current()

        if not input_file or not os.path.exists(input_file):
            messagebox.showwarning("Chưa đủ thông tin", "Vui lòng chọn file Excel nguồn hợp lệ.")
            return

        if not output_path:
            messagebox.showwarning("Chưa đủ thông tin", "Vui lòng nhập đường dẫn nơi lưu file kết quả.")
            return

        if col_idx_sel < 0 or col_idx_sel >= len(self.columns_info):
            messagebox.showwarning("Chưa đủ thông tin", "Vui lòng chọn cột phân nhóm.")
            return

        if not self.selected_group_names:
            messagebox.showwarning("Chưa đủ thông tin", "Bạn chưa chọn nhóm test case nào để xuất sheet.")
            return

        try:
            h_row = int(self.spin_header_row.get() or "1")
        except ValueError:
            h_row = 1

        group_col_idx = self.columns_info[col_idx_sel]["index"]
        forward_fill = self.var_forward_fill.get()
        copy_styles = self.var_styles.get()
        include_summary = self.var_summary.get()
        add_back = self.var_back_link.get()
        export_mode = self.var_export_mode.get()
        selected_keys = list(self.selected_group_names)

        # Khóa nút và cập nhật UI
        self.is_processing = True
        self.btn_run.config(state="disabled", text="⏳ ĐANG XỬ LÝ SINH SHEET...", bg="#94A3B8")
        self.progress_bar["value"] = 0
        self.lbl_percent.config(text="0%")
        self.lbl_status.config(text="Bắt đầu khởi tạo...", fg=COLOR_PRIMARY)

        # Khởi chạy trong Thread riêng biệt
        def worker():
            try:
                def progress_cb(pct: float, msg: str):
                    self.root.after(0, self._update_progress, pct, msg)

                result = generate_excel_sheets(
                    input_file=input_file,
                    output_file=output_path,
                    source_sheet_name=sheet_name,
                    group_column_idx=group_col_idx,
                    header_row=h_row,
                    selected_group_keys=selected_keys,
                    forward_fill=forward_fill,
                    copy_styles=copy_styles,
                    include_summary_sheet=include_summary,
                    add_back_link=add_back,
                    export_mode=export_mode,
                    progress_callback=progress_cb
                )
                self.root.after(0, self._on_generate_success, result)

            except Exception as e:
                self.root.after(0, self._on_generate_error, str(e))

        t = threading.Thread(target=worker, daemon=True)
        t.start()

    def _update_progress(self, percent: float, message: str):
        """Cập nhật thanh tiến trình từ worker thread."""
        self.progress_bar["value"] = percent
        self.lbl_percent.config(text=f"{int(percent)}%")
        self.lbl_status.config(text=message, fg=COLOR_PRIMARY)

    def _on_generate_success(self, result: Dict[str, Any]):
        """Xử lý khi sinh sheet hoàn tất thành công."""
        self.is_processing = False
        self.btn_run.config(state="normal", text="🚀 BẮT ĐẦU SINH SHEET TEST CASE", bg=COLOR_SUCCESS)
        self.progress_bar["value"] = 100
        self.lbl_percent.config(text="100%")
        self.lbl_status.config(text="Đã hoàn tất thành công!", fg=COLOR_SUCCESS)

        output_path = result["output_path"]
        total_groups = result["total_groups"]
        total_testcases = result["total_testcases"]

        self._show_success_dialog(output_path, total_groups, total_testcases)

    def _on_generate_error(self, err_msg: str):
        """Xử lý khi có lỗi xảy ra."""
        self.is_processing = False
        self.btn_run.config(state="normal", text="🚀 BẮT ĐẦU SINH SHEET TEST CASE", bg=COLOR_SUCCESS)
        self.lbl_status.config(text="Có lỗi xảy ra trong quá trình sinh sheet", fg=COLOR_DANGER)
        messagebox.showerror("Lỗi xử lý", f"Đã xảy ra lỗi:\n{err_msg}")

    def _show_success_dialog(self, output_path: str, total_groups: int, total_testcases: int):
        """Hộp thoại thông báo hoàn thành với các nút hành động nhanh."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Sinh Sheet Hoàn Tất")
        dlg.geometry("520x280")
        dlg.resizable(False, False)
        dlg.configure(bg=BG_CARD)
        dlg.transient(self.root)
        dlg.grab_set()

        # Tiêu đề
        top = tk.Frame(dlg, bg=COLOR_SUCCESS, height=50)
        top.pack(fill="x")
        top.pack_propagate(False)

        tk.Label(
            top,
            text="🎉 XUẤT SHEET THÀNH CÔNG!",
            font=(FONT_FAMILY, 13, "bold"),
            bg=COLOR_SUCCESS,
            fg="#FFFFFF"
        ).pack(side="left", padx=16, pady=12)

        # Nội dung
        body = tk.Frame(dlg, bg=BG_CARD, padx=20, pady=16)
        body.pack(fill="both", expand=True)

        msg_text = (
            f"✔ Đã tạo thành công {total_groups} Sheet chi tiết\n"
            f"✔ Tổng số {total_testcases} test case con đã được phân loại chính xác\n\n"
            f"Đường dẫn lưu:\n{output_path}"
        )
        lbl_info = tk.Label(body, text=msg_text, font=FONT_BODY, bg=BG_CARD, fg=TEXT_MAIN, justify="left", wraplength=480)
        lbl_info.pack(anchor="w", pady=(0, 16))

        # Nút bấm hành động
        btn_bar = tk.Frame(body, bg=BG_CARD)
        btn_bar.pack(fill="x")

        btn_open_file = tk.Button(
            btn_bar,
            text="📊 Mở File Excel Ngay",
            font=FONT_BODY_BOLD,
            bg=COLOR_PRIMARY,
            fg="#FFFFFF",
            relief="flat",
            cursor="hand2",
            padx=12,
            pady=6,
            command=lambda: self._open_file_or_dir(output_path)
        )
        btn_open_file.pack(side="left", padx=(0, 8))

        btn_open_folder = tk.Button(
            btn_bar,
            text="📂 Mở Thư Mục Chứa",
            font=FONT_BODY_BOLD,
            bg="#E2E8F0",
            fg=TEXT_MAIN,
            relief="flat",
            cursor="hand2",
            padx=12,
            pady=6,
            command=lambda: self._open_file_folder(output_path)
        )
        btn_open_folder.pack(side="left")

        btn_close = tk.Button(
            btn_bar,
            text="Đóng",
            font=FONT_BODY,
            bg="#CBD5E1",
            fg=TEXT_MAIN,
            relief="flat",
            cursor="hand2",
            padx=16,
            pady=6,
            command=dlg.destroy
        )
        btn_close.pack(side="right")

    def _open_file_or_dir(self, target_path: str):
        """Mở file hoặc thư mục trên Windows."""
        try:
            if sys.platform == "win32":
                os.startfile(target_path)
            else:
                subprocess.Popen(["xdg-open", target_path])
        except Exception as e:
            messagebox.showwarning("Không thể mở", f"Không thể mở trực tiếp: {str(e)}")

    def _open_file_folder(self, target_path: str):
        """Mở thư mục chứa file."""
        try:
            folder = os.path.dirname(os.path.abspath(target_path)) if os.path.isfile(target_path) else target_path
            if sys.platform == "win32":
                subprocess.Popen(["explorer", folder])
            else:
                subprocess.Popen(["xdg-open", folder])
        except Exception as e:
            messagebox.showwarning("Không thể mở thư mục", f"Không thể mở thư mục: {str(e)}")
