#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Label Export Engine (Bộ máy xử lý và định dạng nhãn cáp)
-------------------------------------------------------
Hỗ trợ đa định dạng xuất (Plain text, Tab-delimited Excel, HTML Table cho Word/Docs,
Markdown cho Chat, và trực tiếp tệp tin .xlsx qua openpyxl) nhằm loại bỏ hoàn toàn
vấn đề lệch dòng/lệch cột khi sao chép và dán (Copy & Paste).
"""

import html
import io
import re
import sys
import unicodedata
from typing import List, Tuple

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import pyperclip
    HAS_PYPERCLIP = True
except ImportError:
    HAS_PYPERCLIP = False


def natural_sort_key(s: str):
    """Tách chuỗi thành các phần số và chữ để sắp xếp tự nhiên (1, 2, 10 thay vì 1, 10, 2)."""
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r"(\d+)", s)]


def parse_raw_lines(raw_text: str, auto_sort: bool = False) -> List[Tuple[str, str, str]]:
    """
    Phân tích chuỗi dữ liệu thô thành danh sách (fr_val, idx, to_val).
    Tự động chuẩn hóa Unicode NFC để tránh lệch độ dài ký tự tiếng Việt.
    Nếu auto_sort=True, tự động sắp xếp theo số nhãn tăng dần (natural sorting).
    """
    normalized = unicodedata.normalize("NFC", raw_text)
    cleaned = (
        normalized
        .replace("\xa0", " ")
        .replace("\u200b", "")
        .replace("\u3000", " ")
    )
    lines = cleaned.strip().split("\n")
    
    pattern = re.compile(
        r"Fr:\s*(.+?)\s*\[([^\]]+)\]\s+To:\s*(.+?)(?:\s*\[([^\]]+)\])?$",
        re.IGNORECASE
    )

    parsed = []
    for line in lines:
        line = line.strip()
        if not line:
            continue
        m = pattern.search(line)
        if m:
            fr_val = m.group(1).strip()
            idx = m.group(2).strip()
            to_val = m.group(3).strip()
            parsed.append((fr_val, idx, to_val))

    if auto_sort:
        parsed.sort(key=lambda item: natural_sort_key(item[1]))

    return parsed


def format_plaintext_block(
    fr_val: str,
    to_val: str,
    idx: str,
    col_separator: str = " " * 12,
    align_brackets: bool = True,
    min_padding: int = 2
) -> str:
    """Tạo khối văn bản 4 dòng chuẩn monospace (như logic main.py)."""
    f1 = f"Fr: {fr_val}"
    t1 = f"To: {to_val}"
    f2 = f"Fr: {to_val}"
    t2 = f"To: {fr_val}"

    max_len = max(len(f1), len(t1), len(f2), len(t2))

    def format_row(left_text: str) -> str:
        if align_brackets:
            padding = " " * (max_len - len(left_text) + min_padding)
        else:
            padding = " " * 4
        single_col = f"{left_text}{padding}[{idx}]"
        return f"{single_col}{col_separator}{single_col}"

    return (
        f"{format_row(f1)}\n"
        f"{format_row(t1)}\n\n"
        f"{format_row(f2)}\n"
        f"{format_row(t2)}"
    )


def format_excel_tab_block(fr_val: str, to_val: str, idx: str, separate_id: bool = True) -> str:
    """
    Tạo khối văn bản phân tách bằng ký tự Tab '\\t'.
    Khi dán vào Excel, mỗi đầu cáp hoặc [số] sẽ tự động nhảy vào đúng ô/cột riêng biệt.
    """
    f1 = f"Fr: {fr_val}"
    t1 = f"To: {to_val}"
    f2 = f"Fr: {to_val}"
    t2 = f"To: {fr_val}"

    if separate_id:
        # Cột A: Tên cổng | Cột B: [Số] | Cột C: (trống) | Cột D: Tên cổng | Cột E: [Số]
        return (
            f"{f1}\t[{idx}]\t\t{f1}\t[{idx}]\n"
            f"{t1}\t[{idx}]\t\t{t1}\t[{idx}]\n\n"
            f"{f2}\t[{idx}]\t\t{f2}\t[{idx}]\n"
            f"{t2}\t[{idx}]\t\t{t2}\t[{idx}]"
        )
    else:
        # Cột A: Cổng + [Số] | Cột B: Cổng + [Số]
        return (
            f"{f1}  [{idx}]\t{f1}  [{idx}]\n"
            f"{t1}  [{idx}]\t{t1}  [{idx}]\n\n"
            f"{f2}  [{idx}]\t{f2}  [{idx}]\n"
            f"{t2}  [{idx}]\t{t2}  [{idx}]"
        )


def format_html_table(parsed_list: List[Tuple[str, str, str]]) -> str:
    """
    Tạo biểu diễn bảng HTML Table chuẩn font Consolas/Menlo có đường viền.
    Khi dán vào Word, Excel, Google Docs hay Email, ứng dụng nhận dạng cấu trúc bảng
    và không bao giờ bị xô lệch font chữ hay khoảng trắng.
    """
    rows = []
    for fr_val, idx, to_val in parsed_list:
        f1 = f"Fr: {html.escape(fr_val)}"
        t1 = f"To: {html.escape(to_val)}"
        f2 = f"Fr: {html.escape(to_val)}"
        t2 = f"To: {html.escape(fr_val)}"

        td_style = "padding: 3px 8px; font-family: 'Consolas', 'Menlo', 'Courier New', monospace; font-size: 10pt; border: 1px solid #cbd5e1;"
        td_id = "padding: 3px 8px; font-family: 'Consolas', 'Menlo', monospace; font-size: 10pt; font-weight: bold; text-align: center; color: #1e40af; border: 1px solid #cbd5e1;"
        td_spacer = "width: 24px; border: none;"

        rows.append(f"""
        <tr>
            <td style="{td_style}">{f1}</td>
            <td style="{td_id}">[{idx}]</td>
            <td style="{td_spacer}"></td>
            <td style="{td_style}">{f1}</td>
            <td style="{td_id}">[{idx}]</td>
        </tr>
        <tr>
            <td style="{td_style}">{t1}</td>
            <td style="{td_id}">[{idx}]</td>
            <td style="{td_spacer}"></td>
            <td style="{td_style}">{t1}</td>
            <td style="{td_id}">[{idx}]</td>
        </tr>
        <tr style="height: 10px;"><td colspan="5" style="border: none;"></td></tr>
        <tr>
            <td style="{td_style}">{f2}</td>
            <td style="{td_id}">[{idx}]</td>
            <td style="{td_spacer}"></td>
            <td style="{td_style}">{f2}</td>
            <td style="{td_id}">[{idx}]</td>
        </tr>
        <tr>
            <td style="{td_style}">{t2}</td>
            <td style="{td_id}">[{idx}]</td>
            <td style="{td_spacer}"></td>
            <td style="{td_style}">{t2}</td>
            <td style="{td_id}">[{idx}]</td>
        </tr>
        <tr style="height: 20px;"><td colspan="5" style="border: none;"></td></tr>
        """)

    return f"""<table style="border-collapse: collapse; font-family: 'Consolas', 'Menlo', monospace; font-size: 10pt;">{''.join(rows)}</table>"""


def create_excel_workbook(parsed_list: List[Tuple[str, str, str]]):
    """Tạo đối tượng Workbook openpyxl chứa toàn bộ nhãn cáp đã căn lề sẵn."""
    if not HAS_OPENPYXL:
        return None

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Nhan_Cap_Mang"
    ws.views.sheetView[0].showGridLines = True

    font_main = Font(name="Consolas", size=10, bold=False)
    font_id = Font(name="Consolas", size=10, bold=True, color="1E40AF")
    align_id = Alignment(horizontal="center", vertical="center")

    thin = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )

    row_idx = 1
    for fr_val, idx, to_val in parsed_list:
        f1 = f"Fr: {fr_val}"
        t1 = f"To: {to_val}"
        f2 = f"Fr: {to_val}"
        t2 = f"To: {fr_val}"

        # Dòng 1
        ws.cell(row=row_idx, column=1, value=f1).font = font_main
        ws.cell(row=row_idx, column=1).border = thin
        c2 = ws.cell(row=row_idx, column=2, value=f"[{idx}]")
        c2.font = font_id; c2.alignment = align_id; c2.border = thin
        ws.cell(row=row_idx, column=4, value=f1).font = font_main
        ws.cell(row=row_idx, column=4).border = thin
        c5 = ws.cell(row=row_idx, column=5, value=f"[{idx}]")
        c5.font = font_id; c5.alignment = align_id; c5.border = thin
        row_idx += 1

        # Dòng 2
        ws.cell(row=row_idx, column=1, value=t1).font = font_main
        ws.cell(row=row_idx, column=1).border = thin
        c2 = ws.cell(row=row_idx, column=2, value=f"[{idx}]")
        c2.font = font_id; c2.alignment = align_id; c2.border = thin
        ws.cell(row=row_idx, column=4, value=t1).font = font_main
        ws.cell(row=row_idx, column=4).border = thin
        c5 = ws.cell(row=row_idx, column=5, value=f"[{idx}]")
        c5.font = font_id; c5.alignment = align_id; c5.border = thin
        row_idx += 2

        # Dòng 4 (Nghịch)
        ws.cell(row=row_idx, column=1, value=f2).font = font_main
        ws.cell(row=row_idx, column=1).border = thin
        c2 = ws.cell(row=row_idx, column=2, value=f"[{idx}]")
        c2.font = font_id; c2.alignment = align_id; c2.border = thin
        ws.cell(row=row_idx, column=4, value=f2).font = font_main
        ws.cell(row=row_idx, column=4).border = thin
        c5 = ws.cell(row=row_idx, column=5, value=f"[{idx}]")
        c5.font = font_id; c5.alignment = align_id; c5.border = thin
        row_idx += 1

        # Dòng 5 (Nghịch)
        ws.cell(row=row_idx, column=1, value=t2).font = font_main
        ws.cell(row=row_idx, column=1).border = thin
        c2 = ws.cell(row=row_idx, column=2, value=f"[{idx}]")
        c2.font = font_id; c2.alignment = align_id; c2.border = thin
        ws.cell(row=row_idx, column=4, value=t2).font = font_main
        ws.cell(row=row_idx, column=4).border = thin
        c5 = ws.cell(row=row_idx, column=5, value=f"[{idx}]")
        c5.font = font_id; c5.alignment = align_id; c5.border = thin
        row_idx += 3

    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 4
    ws.column_dimensions["D"].width = 46
    ws.column_dimensions["E"].width = 8

    return wb


def format_html_monospace_pre(formatted_text: str) -> str:
    """
    Bọc chuỗi kết quả trong thẻ <pre> và ép cứng font monospace theo chuẩn:
    <pre style="font-family: 'SF Mono', Consolas, 'Courier New', monospace; font-size: 13px; margin: 0;">...
    """
    escaped = html.escape(formatted_text)
    return f"""<pre style="font-family: 'SF Mono', Consolas, 'Courier New', monospace; font-size: 13px; margin: 0;">{escaped}</pre>"""


def copy_multi_format_to_clipboard(plain_text: str, html_text: str = None) -> bool:
    """
    Sao chép đồng thời cả dạng Văn bản thuần (Plain Text) và Khối monospace ép font (HTML <pre>)
    vào Clipboard hệ điều hành.
    - Khi dán vào Text Editor / Notepad / In nhãn: Nhận dạng văn bản thuần monospace.
    - Khi dán vào Word / Docs / Pages / Email: Nhận dạng HTML <pre> ép font SF Mono/Consolas,
      giữ nguyên từng khoảng trắng và vị trí cột không bao giờ bị lệch!
    """
    if not html_text:
        html_text = format_html_monospace_pre(plain_text)

    success = False

    # 1. macOS Native AppKit NSPasteboard (Hỗ trợ đa định dạng đồng thời)
    if sys.platform == "darwin":
        try:
            from AppKit import NSPasteboard, NSPasteboardTypeString, NSPasteboardTypeHTML
            pb = NSPasteboard.generalPasteboard()
            pb.clearContents()
            pb.setString_forType_(plain_text, NSPasteboardTypeString)
            pb.setString_forType_(html_text, NSPasteboardTypeHTML)
            success = True
        except Exception:
            pass

    # 2. Fallback pyperclip nếu không phải macOS hoặc AppKit lỗi
    if not success and HAS_PYPERCLIP:
        try:
            pyperclip.copy(plain_text)
            success = True
        except Exception:
            pass

    return success
