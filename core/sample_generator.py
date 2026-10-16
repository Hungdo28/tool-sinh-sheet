# -*- coding: utf-8 -*-
"""
Sample Generator Module
Tạo file Excel Test Case tổng chuẩn 100% theo mẫu của người dùng:
- Header 2 tầng màu xanh Cyan (#00B0F0 / #00CCFF)
- Gộp các cột: Data test (Lần 1, 2, 3), UAT (Lần 1, 2, 3), PROD (Lần 1, 2, 3)
- Cột '(ID file log bug/ Jira)' chữ đỏ nghiêng
- Dòng banner màu vàng phân loại Job / Module
- Dòng test case con với ID: [Tên_TC]_1, [Tên_TC]_2...
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def create_sample_master_file(output_path: str = "sample_data/sample_master_testcases.xlsx") -> str:
    """Tạo file mẫu chuẩn 100% khớp với ảnh người dùng cung cấp."""
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Master_TestCases"

    # ================= STYLES =================
    # Màu xanh Cyan Header (#00B0F0) - Bắt buộc có prefix FF (Alpha = 100% Opaque) để Excel hiển thị màu
    cyan_fill = PatternFill(start_color="FF00B0F0", end_color="FF00B0F0", fill_type="solid")
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FF000000")
    header_red_font = Font(name="Segoe UI", size=10, bold=True, italic=True, color="FFFF0000")

    # Màu vàng Banner Job (#FFE699) - Prefix FF để Excel đổ màu vàng đặc
    yellow_fill = PatternFill(start_color="FFFFE699", end_color="FFFFE699", fill_type="solid")
    yellow_font = Font(name="Segoe UI", size=10, bold=True, color="FF000000")

    # Border mỏng màu đen (Prefix FF)
    thin_side = Side(style='thin', color='FF000000')
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    data_font = Font(name="Segoe UI", size=9.5)

    # ================= HEADER ROW 1 & ROW 2 =================
    # Danh sách 20 cột
    # Col 1: ID
    # Col 2: Mục đích kiểm thử
    # Col 3: Các bước thực hiện
    # Col 4: Kết quả mong muốn
    # Col 5: thực tế (UAT)
    # Col 6: thực tế (Prod)
    # Col 7-9: Data test (Lần 1, 2, 3)
    # Col 10-12: UAT (Lần 1, 2, 3)
    # Col 13-15: PROD (Lần 1, 2, 3)
    # Col 16: Kết quả hiện tại
    # Col 17: (ID file log bug/ Jira)
    # Col 18: Ghi chú
    # Col 19: Ngày test UAT
    # Col 20: Ngày test PROD

    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 25

    # Tô màu và viền cho tất cả các ô trong 2 dòng header trước
    for r in (1, 2):
        for c in range(1, 21):
            cell = ws.cell(row=r, column=c)
            cell.fill = cyan_fill
            cell.border = cell_border
            cell.font = header_font
            cell.alignment = center_align

    # Điền giá trị Header tầng 1
    ws.cell(row=1, column=1, value="ID")
    ws.cell(row=1, column=2, value="Mục đích kiểm thử")
    ws.cell(row=1, column=3, value="Các bước thực hiện")
    ws.cell(row=1, column=4, value="Kết quả mong muốn")
    ws.cell(row=1, column=5, value="thực tế (UAT)")
    ws.cell(row=1, column=6, value="thực tế (Prod)")
    ws.cell(row=1, column=7, value="Data test")
    ws.cell(row=1, column=10, value="UAT")
    ws.cell(row=1, column=13, value="PROD")
    ws.cell(row=1, column=16, value="Kết quả hiện tại")
    cell_jira = ws.cell(row=1, column=17, value="(ID file log bug/ Jira)")
    cell_jira.font = header_red_font
    ws.cell(row=1, column=18, value="Ghi chú")
    ws.cell(row=1, column=19, value="Ngày test UAT")
    ws.cell(row=1, column=20, value="Ngày test PROD")

    # Điền giá trị Header tầng 2
    ws.cell(row=2, column=7, value="Lần 1")
    ws.cell(row=2, column=8, value="Lần 2")
    ws.cell(row=2, column=9, value="Lần 3")
    ws.cell(row=2, column=10, value="Lần 1")
    ws.cell(row=2, column=11, value="Lần 2")
    ws.cell(row=2, column=12, value="Lần 3")
    ws.cell(row=2, column=13, value="Lần 1")
    ws.cell(row=2, column=14, value="Lần 2")
    ws.cell(row=2, column=15, value="Lần 3")

    # Gộp ô cho các cột không chia tầng (gộp hàng 1 & 2)
    single_cols = [1, 2, 3, 4, 5, 6, 16, 17, 18, 19, 20]
    for col in single_cols:
        ws.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col)

    # Gộp ô tầng 1 cho Data test, UAT, PROD
    ws.merge_cells(start_row=1, start_column=7, end_row=1, end_column=9)    # Data test
    ws.merge_cells(start_row=1, start_column=10, end_row=1, end_column=12) # UAT
    ws.merge_cells(start_row=1, start_column=13, end_row=1, end_column=15) # PROD

    # ================= ROW 3: YELLOW BANNER =================
    ws.row_dimensions[3].height = 24
    for c in range(1, 21):
        cell = ws.cell(row=3, column=c)
        cell.fill = yellow_fill
        cell.border = cell_border
    ws.cell(row=3, column=2, value="Job quét và tổng hợp danh sách nợ").font = yellow_font
    ws.cell(row=3, column=2).alignment = left_align

    # ================= ROWS 4+: TEST CASES CON =================
    testcases_data = [
        (
            "[Tên_TC]_1",
            "Xác minh hệ thống quét bản ghi thỏa mãn điều kiện [STATUS != 1, FAILURE_REASON = 1, IS_SCAN > 0]",
            "1. Bảng SETTLEMENT.PRE_TRANS có bản ghi với STATUS=2, FAILURE_REASON=1, IS_SCAN=1",
            "2. - Bản ghi được chọn để xử lý (được đánh dấu hoặc chuyển trạng thái theo quy trình nội bộ)"
        ),
        (
            "[Tên_TC]_2",
            "Xác minh hệ thống bỏ qua bản ghi không thỏa mãn điều kiện quét [STATUS = 1]",
            "1. Bảng SETTLEMENT.PRE_TRANS có bản ghi với STATUS=1, FAILURE_REASON=1, IS_SCAN=1",
            "2. - Bản ghi bị bỏ qua, không được chọn đưa vào danh sách xử lý nợ"
        ),
        (
            "[Tên_TC]_3",
            "Xác minh hệ thống bỏ qua bản ghi có FAILURE_REASON khác 1",
            "1. Bảng SETTLEMENT.PRE_TRANS có bản ghi với STATUS=2, FAILURE_REASON=0, IS_SCAN=1",
            "2. - Hệ thống không quét bản ghi này"
        ),
        (
            "[Tên_TC]_4",
            "Xác minh hệ thống xử lý khi IS_SCAN = 0",
            "1. Bảng SETTLEMENT.PRE_TRANS có bản ghi với STATUS=2, FAILURE_REASON=1, IS_SCAN=0",
            "2. - Bản ghi không thỏa mãn điều kiện IS_SCAN > 0, bị bỏ qua"
        ),
        (
            "[Tên_TC]_5",
            "Kiểm tra hiệu năng khi bảng SETTLEMENT.PRE_TRANS có 100,000 bản ghi",
            "1. Chuẩn bị 100k bản ghi thỏa mãn điều kiện quét\n2. Kích hoạt Job quét danh sách nợ",
            "2. - Thời gian thực thi Job < 5 phút, không bị timeout hay khóa bảng"
        )
    ]

    start_row = 4
    for idx, (tc_id, tc_desc, tc_steps, tc_expected) in enumerate(testcases_data):
        curr_r = start_row + idx
        ws.row_dimensions[curr_r].height = 65

        for c in range(1, 21):
            cell = ws.cell(row=curr_r, column=c)
            cell.font = data_font
            cell.border = cell_border
            cell.alignment = left_align

        ws.cell(row=curr_r, column=1, value=tc_id).alignment = center_align
        ws.cell(row=curr_r, column=2, value=tc_desc)
        ws.cell(row=curr_r, column=3, value=tc_steps)
        ws.cell(row=curr_r, column=4, value=tc_expected)

    # ================= COLUMN WIDTHS =================
    col_widths = {
        'A': 16,  # ID
        'B': 38,  # Mục đích kiểm thử
        'C': 42,  # Các bước thực hiện
        'D': 42,  # Kết quả mong muốn
        'E': 14,  # thực tế (UAT)
        'F': 14,  # thực tế (Prod)
        'G': 10,  # Data test Lần 1
        'H': 10,  # Data test Lần 2
        'I': 10,  # Data test Lần 3
        'J': 10,  # UAT Lần 1
        'K': 10,  # UAT Lần 2
        'L': 10,  # UAT Lần 3
        'M': 10,  # PROD Lần 1
        'N': 10,  # PROD Lần 2
        'O': 10,  # PROD Lần 3
        'P': 16,  # Kết quả hiện tại
        'Q': 24,  # (ID file log bug/ Jira)
        'R': 18,  # Ghi chú
        'S': 16,  # Ngày test UAT
        'T': 16   # Ngày test PROD
    }

    for letter, width in col_widths.items():
        ws.column_dimensions[letter].width = width

    wb.save(output_path)
    wb.close()
    return output_path


if __name__ == "__main__":
    path = create_sample_master_file()
    print("Created exact format sample file:", path)
