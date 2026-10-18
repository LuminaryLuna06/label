#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cable Label Generator GUI (Giao Diện Tạo Nhãn Cáp Mạng)
------------------------------------------------------
Giao diện trực quan hỗ trợ xử lý dữ liệu cổng/thiết bị từ Excel thành
danh sách nhãn cáp 2 chiều chuẩn hóa, kèm nút sao chép nhanh (copy to clipboard).
"""

import os
import re
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import unicodedata

import label_export

try:
    import pyperclip
    HAS_PYPERCLIP = True
except ImportError:
    HAS_PYPERCLIP = False

# Font gia đình chuẩn cho giao diện và văn bản đơn cách (monospace)
UI_FONT = ("SF Pro Display", 11) if sys.platform == "darwin" else ("Segoe UI", 10)
UI_FONT_BOLD = ("SF Pro Display", 11, "bold") if sys.platform == "darwin" else ("Segoe UI", 10, "bold")
TITLE_FONT = ("SF Pro Display", 15, "bold") if sys.platform == "darwin" else ("Segoe UI", 14, "bold")
MONO_FONT = ("Menlo", 10) if sys.platform == "darwin" else ("Consolas", 10)
SMALL_FONT = ("SF Pro Text", 9) if sys.platform == "darwin" else ("Segoe UI", 8)

# Bảng màu hiện đại
BG_COLOR = "#F1F5F9"        # Nền chính xám nhẹ
PANEL_BG = "#FFFFFF"        # Nền thẻ trắng tinh
PRIMARY_COLOR = "#2563EB"   # Xanh dương chủ đạo
PRIMARY_HOVER = "#1D4ED8"
BORDER_COLOR = "#E2E8F0"
TEXT_MAIN = "#0F172A"
TEXT_MUTED = "#64748B"
SUCCESS_COLOR = "#16A34A"   # Xanh lá khi đã copy thành công
SUCCESS_BG = "#DCFCE7"

SAMPLE_DATA = """Fr: S4_U13-14_SRV số 1_Slot 4_Port 1    [1]\tTo: S4_U41_ODF_port 1->N2_U13_ODF_port 1    [1]
Fr: S4_U13-14_SRV số 1_Slot 4_Port 2    [2]\tTo: S4_U41_ODF_port 2->N2_U13_ODF_port 2    [2]
Fr: S4_U13-14_SRV số 1_iDrac    [3]\tTo: S4_U40_PP_Port 5->N1_U3_PP_Port 5    [3]
Fr: N2_U13_ODF_port 1    [4]\tTo: N3_U1-13_CoreSw01_Slot 10_port 1    [4]
Fr: N2_U13_ODF_port 2    [5]\tTo: N4_U1-13_CoreSw02_Slot 10_port 1    [5]
Fr: N1_U3_PP_Port 5    [6]\tTo: N1_U20_SWMgmt_port 18    [6]"""


def parse_and_generate_labels(
    raw_text: str,
    col_separator: str = " " * 12,
    align_brackets: bool = True,
    min_bracket_padding: int = 2,
    auto_sort: bool = True
):
    """
    Phân tích chuỗi đầu vào theo cú pháp Fr: ... [id] To: ... [id]
    và trả về danh sách các tuple: (label_id, formatted_block_text).
    Tự động chuẩn hóa Unicode NFC để đảm bảo độ dài ký tự tiếng Việt chính xác.
    Hỗ trợ tự động sắp xếp (auto_sort=True) theo thứ tự nhãn tăng dần tự nhiên (1, 2, 10...).
    """
    normalized_text = unicodedata.normalize("NFC", raw_text)
    cleaned_text = (
        normalized_text
        .replace("\xa0", " ")
        .replace("\u200b", "")
        .replace("\u3000", " ")
    )
    lines = cleaned_text.strip().split("\n")
    blocks = []

    # Bắt định dạng Fr: <nội dung> [số/kí tự] To: <nội dung> [số/kí tự]
    pattern = re.compile(
        r"Fr:\s*(.+?)\s*\[([^\]]+)\]\s+To:\s*(.+?)(?:\s*\[([^\]]+)\])?$",
        re.IGNORECASE
    )

    for line in lines:
        line = line.strip()
        if not line:
            continue

        match = pattern.search(line)
        if match:
            fr_val = match.group(1).strip()
            idx = match.group(2).strip()
            to_val = match.group(3).strip()

            # Chiều thuận (Đầu 1)
            f1 = f"Fr: {fr_val}"
            t1 = f"To: {to_val}"

            # Chiều nghịch (Đầu 2 - Đảo chiều)
            f2 = f"Fr: {to_val}"
            t2 = f"To: {fr_val}"

            # Độ dài lớn nhất để căn chỉnh cột chuẩn xác
            max_len = max(len(f1), len(t1), len(f2), len(t2))

            def format_row(left_text: str, label_id: str) -> str:
                if align_brackets:
                    padding = " " * (max_len - len(left_text) + min_bracket_padding)
                else:
                    padding = " " * 4
                single_col = f"{left_text}{padding}[{label_id}]"
                return f"{single_col}{col_separator}{single_col}"

            block = (
                f"{format_row(f1, idx)}\n"
                f"{format_row(t1, idx)}\n\n"
                f"{format_row(f2, idx)}\n"
                f"{format_row(t2, idx)}"
            )
            blocks.append((idx, block))

    if auto_sort:
        blocks.sort(key=lambda b: label_export.natural_sort_key(b[0]))

    return blocks


class CableLabelApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tạo Nhãn Cáp Mạng • Cable Label Tool")
        self.geometry("1180x760")
        self.minsize(900, 600)
        self.configure(bg=BG_COLOR)

        self.current_labels = []  # Danh sách lưu (id, block_text)
        self.current_parsed = []  # Danh sách lưu (fr, id, to)
        self.copy_timer_handles = {}  # Lưu id hẹn giờ reset nút copy

        self._init_style()
        self._build_header()
        self._build_main_panes()
        self._build_status_bar()

    def _init_style(self):
        self.style = ttk.Style(self)
        # Sử dụng theme 'clam' hoặc 'aqua'
        available_themes = self.style.theme_names()
        if "clam" in available_themes:
            self.style.theme_use("clam")

        self.style.configure(".", font=UI_FONT)
        self.style.configure("TFrame", background=BG_COLOR)
        self.style.configure("Card.TFrame", background=PANEL_BG)
        self.style.configure("TLabel", background=PANEL_BG, foreground=TEXT_MAIN, font=UI_FONT)

    def _build_header(self):
        header_frame = tk.Frame(self, bg=PANEL_BG, padx=20, pady=12, highlightthickness=1, highlightbackground=BORDER_COLOR)
        header_frame.pack(fill=tk.X, side=tk.TOP)

        title_label = tk.Label(
            header_frame,
            text="🏷️ Công Cụ Tạo Nhãn Cáp Mạng",
            font=TITLE_FONT,
            bg=PANEL_BG,
            fg=PRIMARY_COLOR
        )
        title_label.pack(anchor=tk.W)

        subtitle_label = tk.Label(
            header_frame,
            text="Chuyển đổi dữ liệu cổng từ Excel thành nhãn cáp 2 chiều chuẩn hóa • Nhấn Copy để sao chép nhanh vào Clipboard",
            font=SMALL_FONT,
            bg=PANEL_BG,
            fg=TEXT_MUTED
        )
        subtitle_label.pack(anchor=tk.W, pady=(2, 0))

    def _build_main_panes(self):
        container = tk.Frame(self, bg=BG_COLOR, padx=16, pady=14)
        container.pack(fill=tk.BOTH, expand=True)

        # Sử dụng PanedWindow để người dùng có thể kéo co giãn giữa 2 bên
        self.paned = ttk.PanedWindow(container, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True)

        # --- KHỐI TRÁI: INPUT ---
        left_frame = tk.Frame(self.paned, bg=PANEL_BG, padx=14, pady=14, highlightthickness=1, highlightbackground=BORDER_COLOR)
        self.paned.add(left_frame, weight=1)

        left_header = tk.Frame(left_frame, bg=PANEL_BG)
        left_header.pack(fill=tk.X, pady=(0, 8))

        lbl_input = tk.Label(left_header, text="📥 Dữ liệu đầu vào", font=UI_FONT_BOLD, bg=PANEL_BG, fg=TEXT_MAIN)
        lbl_input.pack(side=tk.LEFT)

        btn_clear = tk.Button(
            left_header, text="Xóa", font=SMALL_FONT, bg="#F1F5F9", fg=TEXT_MUTED,
            relief="flat", cursor="hand2", padx=8, pady=2, command=self.clear_input
        )
        btn_clear.pack(side=tk.RIGHT, padx=(4, 0))

        btn_sample = tk.Button(
            left_header, text="Dán mẫu", font=SMALL_FONT, bg="#F1F5F9", fg=PRIMARY_COLOR,
            relief="flat", cursor="hand2", padx=8, pady=2, command=self.paste_sample
        )
        btn_sample.pack(side=tk.RIGHT, padx=(4, 0))

        btn_paste = tk.Button(
            left_header, text="Dán từ Clipboard", font=SMALL_FONT, bg="#EFF6FF", fg=PRIMARY_COLOR,
            relief="flat", cursor="hand2", padx=8, pady=2, command=self.paste_clipboard
        )
        btn_paste.pack(side=tk.RIGHT)

        # Ô nhập văn bản
        text_container = tk.Frame(left_frame, bg=BORDER_COLOR, bd=1)
        text_container.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.input_text = tk.Text(
            text_container,
            wrap=tk.NONE,
            font=MONO_FONT,
            padx=10,
            pady=10,
            relief="flat",
            bg="#FAFAFA",
            fg=TEXT_MAIN,
            insertbackground=PRIMARY_COLOR
        )
        scroll_y = ttk.Scrollbar(text_container, orient=tk.VERTICAL, command=self.input_text.yview)
        scroll_x = ttk.Scrollbar(text_container, orient=tk.HORIZONTAL, command=self.input_text.xview)
        self.input_text.configure(xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set)

        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.input_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Tùy chọn phân cách cột
        sep_frame = tk.Frame(left_frame, bg=PANEL_BG)
        sep_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(sep_frame, text="Khoảng cách 2 cột:", font=SMALL_FONT, bg=PANEL_BG, fg=TEXT_MUTED).pack(side=tk.LEFT)

        self.sep_var = tk.StringVar(value="12_spaces")
        rb_spaces = tk.Radiobutton(
            sep_frame, text="12 Spaces (In ấn/Txt)", variable=self.sep_var, value="12_spaces",
            font=SMALL_FONT, bg=PANEL_BG, activebackground=PANEL_BG, highlightthickness=0
        )
        rb_spaces.pack(side=tk.LEFT, padx=(6, 0))

        rb_tab = tk.Radiobutton(
            sep_frame, text="Tab '\\t' (Excel 2 cột)", variable=self.sep_var, value="tab",
            font=SMALL_FONT, bg=PANEL_BG, activebackground=PANEL_BG, highlightthickness=0
        )
        rb_tab.pack(side=tk.LEFT, padx=(6, 0))

        # Tùy chọn căn lề [số]
        align_frame = tk.Frame(left_frame, bg=PANEL_BG)
        align_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(align_frame, text="Căn lề [số]:", font=SMALL_FONT, bg=PANEL_BG, fg=TEXT_MUTED).pack(side=tk.LEFT)

        self.align_var = tk.StringVar(value="dynamic")
        rb_dyn = tk.Radiobutton(
            align_frame, text="Căn thẳng cột (main.py)", variable=self.align_var, value="dynamic",
            font=SMALL_FONT, bg=PANEL_BG, activebackground=PANEL_BG, highlightthickness=0
        )
        rb_dyn.pack(side=tk.LEFT, padx=(6, 0))

        rb_fix = tk.Radiobutton(
            align_frame, text="Cố định 4 spaces", variable=self.align_var, value="fixed",
            font=SMALL_FONT, bg=PANEL_BG, activebackground=PANEL_BG, highlightthickness=0
        )
        rb_fix.pack(side=tk.LEFT, padx=(6, 0))

        # Tùy chọn tự động sắp xếp
        sort_frame = tk.Frame(left_frame, bg=PANEL_BG)
        sort_frame.pack(fill=tk.X, pady=(0, 10))

        self.sort_var = tk.BooleanVar(value=True)
        cb_sort = tk.Checkbutton(
            sort_frame, text="🔢 Tự động sắp xếp nhãn tăng dần (1 → n)",
            variable=self.sort_var, font=SMALL_FONT, bg=PANEL_BG,
            activebackground=PANEL_BG, highlightthickness=0
        )
        cb_sort.pack(side=tk.LEFT)

        # Nút chuyển đổi chính
        self.btn_generate = tk.Button(
            left_frame,
            text="⚡ Tạo Danh Sách Nhãn",
            font=UI_FONT_BOLD,
            bg=PRIMARY_COLOR,
            fg="white",
            activebackground=PRIMARY_HOVER,
            activeforeground="white",
            relief="flat",
            cursor="hand2",
            pady=10,
            command=self.process_generation
        )
        self.btn_generate.pack(fill=tk.X)

        # --- KHỐI PHẢI: OUTPUT ---
        right_frame = tk.Frame(self.paned, bg=PANEL_BG, padx=14, pady=14, highlightthickness=1, highlightbackground=BORDER_COLOR)
        self.paned.add(right_frame, weight=2)

        # Header của Output
        right_header = tk.Frame(right_frame, bg=PANEL_BG)
        right_header.pack(fill=tk.X, pady=(0, 8))

        lbl_output_title = tk.Label(right_header, text="📤 Danh sách nhãn", font=UI_FONT_BOLD, bg=PANEL_BG, fg=TEXT_MAIN)
        lbl_output_title.pack(side=tk.LEFT)

        self.lbl_count = tk.Label(
            right_header,
            text="0 nhãn",
            font=SMALL_FONT,
            bg="#EEF2F6",
            fg=TEXT_MUTED,
            padx=8,
            pady=2
        )
        self.lbl_count.pack(side=tk.LEFT, padx=(8, 0))

        # Nút xuất file .txt
        self.btn_save_file = tk.Button(
            right_header,
            text="💾 Lưu .txt",
            font=SMALL_FONT,
            bg="#F1F5F9",
            fg=TEXT_MAIN,
            relief="flat",
            cursor="hand2",
            padx=10,
            pady=3,
            command=self.save_to_file,
            state=tk.DISABLED
        )
        self.btn_save_file.pack(side=tk.RIGHT, padx=(6, 0))

        # Nút xuất file Excel .xlsx
        self.btn_save_excel = tk.Button(
            right_header,
            text="📊 Xuất Excel (.xlsx)",
            font=SMALL_FONT,
            bg="#ECFDF5",
            fg="#047857",
            relief="flat",
            cursor="hand2",
            padx=10,
            pady=3,
            command=self.save_excel_file,
            state=tk.DISABLED
        )
        self.btn_save_excel.pack(side=tk.RIGHT, padx=(6, 0))

        # Nút Copy Tất Cả
        self.btn_copy_all = tk.Button(
            right_header,
            text="📋 Copy tất cả",
            font=SMALL_FONT,
            bg="#EFF6FF",
            fg=PRIMARY_COLOR,
            relief="flat",
            cursor="hand2",
            padx=10,
            pady=3,
            command=self.copy_all_labels,
            state=tk.DISABLED
        )
        self.btn_copy_all.pack(side=tk.RIGHT)

        # Khung hiển thị danh sách dạng cuộn (Scrollable Canvas)
        self.canvas_frame = tk.Frame(right_frame, bg=BG_COLOR)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(self.canvas_frame, bg=BG_COLOR, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollable_content = tk.Frame(self.canvas, bg=BG_COLOR)

        self.scrollable_content.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_content, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Đảm bảo nội dung vừa khít chiều ngang của canvas khi co giãn cửa sổ
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.itemconfig(self.canvas_window, width=e.width)
        )

        # Cuộn chuột mượt mà trên mọi nền tảng
        self._bind_mousewheel(self.canvas)
        self._bind_mousewheel(self.scrollable_content)

        # Placeholder ban đầu
        self.show_empty_placeholder()

    def _build_status_bar(self):
        self.status_bar = tk.Frame(self, bg=PANEL_BG, padx=16, pady=4, highlightthickness=1, highlightbackground=BORDER_COLOR)
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)

        self.status_label = tk.Label(
            self.status_bar,
            text="Sẵn sàng • Hãy dán dữ liệu vào ô bên trái rồi nhấn 'Tạo Danh Sách Nhãn'",
            font=SMALL_FONT,
            bg=PANEL_BG,
            fg=TEXT_MUTED
        )
        self.status_label.pack(side=tk.LEFT)

    def _bind_mousewheel(self, widget):
        widget.bind("<MouseWheel>", self._on_mousewheel)
        widget.bind("<Button-4>", self._on_mousewheel)
        widget.bind("<Button-5>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        if not self.current_labels:
            return
        if sys.platform == "darwin":
            self.canvas.yview_scroll(int(-1 * event.delta), "units")
        elif sys.platform.startswith("win"):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        else:
            if event.num == 5:
                self.canvas.yview_scroll(1, "units")
            elif event.num == 4:
                self.canvas.yview_scroll(-1, "units")

    def show_empty_placeholder(self):
        for w in self.scrollable_content.winfo_children():
            w.destroy()

        placeholder = tk.Frame(self.scrollable_content, bg=BG_COLOR, pady=60)
        placeholder.pack(fill=tk.BOTH, expand=True)

        lbl_icon = tk.Label(placeholder, text="🏷️", font=("SF Pro Display", 36), bg=BG_COLOR)
        lbl_icon.pack()

        lbl_text = tk.Label(
            placeholder,
            text="Chưa có nhãn nào được tạo.\nDán dữ liệu từ file Excel vào ô bên trái để bắt đầu.",
            font=UI_FONT,
            bg=BG_COLOR,
            fg=TEXT_MUTED,
            justify=tk.CENTER
        )
        lbl_text.pack(pady=10)

        self._bind_mousewheel(placeholder)
        self._bind_mousewheel(lbl_icon)
        self._bind_mousewheel(lbl_text)

    def paste_clipboard(self):
        try:
            content = self.clipboard_get()
            self.input_text.insert(tk.INSERT, content)
            self.status_label.config(text="Đã dán dữ liệu từ Clipboard.")
        except tk.TclError:
            messagebox.showwarning("Clipboard trống", "Không tìm thấy nội dung dạng văn bản trong Clipboard!")

    def paste_sample(self):
        self.input_text.delete("1.0", tk.END)
        self.input_text.insert(tk.END, SAMPLE_DATA)
        self.status_label.config(text="Đã nạp dữ liệu mẫu thử.")

    def clear_input(self):
        self.input_text.delete("1.0", tk.END)
        self.status_label.config(text="Đã xóa ô nhập liệu.")

    def process_generation(self):
        raw_text = self.input_text.get("1.0", tk.END).strip()
        if not raw_text:
            messagebox.showinfo("Thông báo", "Vui lòng nhập hoặc dán dữ liệu trước khi bấm tạo nhãn!")
            return

        auto_sort = self.sort_var.get()
        parsed_items = label_export.parse_raw_lines(raw_text, auto_sort=auto_sort)
        self.current_parsed = parsed_items

        col_sep = "\t" if self.sep_var.get() == "tab" else (" " * 12)
        align_brackets = (self.align_var.get() == "dynamic")
        blocks = parse_and_generate_labels(
            raw_text,
            col_separator=col_sep,
            align_brackets=align_brackets,
            auto_sort=auto_sort
        )
        self.current_labels = blocks

        # Cập nhật số lượng
        count = len(blocks)
        self.lbl_count.config(text=f"{count} nhãn", bg="#DCFCE7" if count > 0 else "#EEF2F6", fg="#15803D" if count > 0 else TEXT_MUTED)

        if count == 0:
            messagebox.showwarning(
                "Không nhận diện được dữ liệu",
                "Không tìm thấy dòng nào khớp định dạng:\nFr: ... [số] To: ... [số]\n\nVui lòng kiểm tra lại cấu trúc dữ liệu đầu vào."
            )
            self.show_empty_placeholder()
            self.btn_copy_all.config(state=tk.DISABLED)
            self.btn_save_file.config(state=tk.DISABLED)
            self.btn_save_excel.config(state=tk.DISABLED)
            self.status_label.config(text="Không tìm thấy dòng nhãn hợp lệ.")
            return

        self.btn_copy_all.config(state=tk.NORMAL)
        self.btn_save_file.config(state=tk.NORMAL)
        self.btn_save_excel.config(state=tk.NORMAL)
        self.status_label.config(text=f"✔ Đã tạo thành công {count} bộ nhãn cáp mạng!")

        # Render danh sách card
        self._render_label_cards()

    def _render_label_cards(self):
        for w in self.scrollable_content.winfo_children():
            w.destroy()

        for (fr_val, idx, to_val), (_, block_text) in zip(self.current_parsed, self.current_labels):
            self._create_label_card(idx, block_text, (fr_val, idx, to_val))

        # Cập nhật cuộn lên đầu trang
        self.canvas.yview_moveto(0)

    def _create_label_card(self, label_id: str, block_text: str, item_tuple: tuple = None):
        card = tk.Frame(
            self.scrollable_content,
            bg=PANEL_BG,
            padx=12,
            pady=10,
            highlightthickness=1,
            highlightbackground=BORDER_COLOR
        )
        card.pack(fill=tk.X, padx=4, pady=6)

        # Header của từng Card
        card_header = tk.Frame(card, bg=PANEL_BG)
        card_header.pack(fill=tk.X, pady=(0, 6))

        tag = tk.Label(
            card_header,
            text=f"🏷️ Nhãn [{label_id}]",
            font=UI_FONT_BOLD,
            bg=PANEL_BG,
            fg=TEXT_MAIN
        )
        tag.pack(side=tk.LEFT)

        # Nút Copy từng nhãn
        btn_copy = tk.Button(
            card_header,
            text="📋 Copy",
            font=SMALL_FONT,
            bg="#EFF6FF",
            fg=PRIMARY_COLOR,
            activebackground=PRIMARY_COLOR,
            activeforeground="white",
            relief="flat",
            cursor="hand2",
            padx=10,
            pady=3
        )
        # Sử dụng lambda với tham chiếu cụ thể
        btn_copy.config(command=lambda b=btn_copy, t=block_text, lid=label_id, it=item_tuple: self.copy_single_label(b, t, lid, it))
        btn_copy.pack(side=tk.RIGHT)

        # Khối văn bản hiển thị nội dung 4 dòng
        code_box = tk.Text(
            card,
            wrap=tk.NONE,
            font=MONO_FONT,
            bg="#F8FAFC",
            fg="#1E293B",
            height=5,
            relief="flat",
            padx=8,
            pady=6,
            highlightthickness=1,
            highlightbackground="#E2E8F0"
        )
        code_box.insert(tk.END, block_text)
        code_box.configure(state=tk.DISABLED)
        code_box.pack(fill=tk.X)

        # Gắn sự kiện cuộn chuột cho toàn bộ widget con của card
        for w in (card, card_header, tag, code_box):
            self._bind_mousewheel(w)

    def copy_to_system_clipboard(self, text: str):
        """
        Sao chép chuỗi chính xác 100% khoảng trắng (spaces/tabs/newlines).
        Ưu tiên dùng pyperclip (gọi trực tiếp pbcopy trên macOS / API OS chuẩn),
        đồng thời đồng bộ vào Tkinter clipboard.
        """
        if HAS_PYPERCLIP:
            try:
                pyperclip.copy(text)
            except Exception:
                pass
        try:
            self.clipboard_clear()
            self.clipboard_append(text)
            self.update()
        except Exception:
            pass

    def copy_single_label(self, button: tk.Button, text_to_copy: str, label_id: str, item_tuple: tuple = None):
        """Sao chép nhãn cụ thể đa định dạng (Plain text + HTML Table) và kích hoạt trạng thái 'Đã copy'."""
        html_table = label_export.format_html_table([item_tuple]) if item_tuple else None
        label_export.copy_multi_format_to_clipboard(text_to_copy, html_table)
        try:
            self.clipboard_clear()
            self.clipboard_append(text_to_copy)
            self.update()
        except Exception:
            pass

        # Cập nhật diện mạo nút sang "Đã copy"
        original_text = "📋 Copy"
        original_bg = "#EFF6FF"
        original_fg = PRIMARY_COLOR

        button.config(
            text="✔ Đã copy",
            bg=SUCCESS_BG,
            fg=SUCCESS_COLOR
        )

        self.status_label.config(text=f"✔ Đã copy nội dung Nhãn [{label_id}] (Đa định dạng chống lệch khi dán vào Excel/Word)!")

        # Hủy lịch hẹn cũ nếu có
        btn_id = id(button)
        if btn_id in self.copy_timer_handles:
            self.after_cancel(self.copy_timer_handles[btn_id])

        # Khôi phục nút sau 1.8 giây
        def reset_button():
            try:
                button.config(
                    text=original_text,
                    bg=original_bg,
                    fg=original_fg
                )
            except tk.TclError:
                pass  # Trong trường hợp widget đã bị hủy

        self.copy_timer_handles[btn_id] = self.after(1800, reset_button)

    def copy_all_labels(self):
        """Sao chép toàn bộ nhãn cáp vào Clipboard đa định dạng (Plain text + HTML Table)."""
        if not self.current_labels:
            return

        all_content = "\n\n\n".join(b for _, b in self.current_labels)
        html_table = label_export.format_html_table(self.current_parsed) if self.current_parsed else None
        label_export.copy_multi_format_to_clipboard(all_content, html_table)
        try:
            self.clipboard_clear()
            self.clipboard_append(all_content)
            self.update()
        except Exception:
            pass

        # Hiệu ứng đổi chữ nút "Copy tất cả"
        self.btn_copy_all.config(
            text="✔ Đã copy tất cả",
            bg=SUCCESS_BG,
            fg=SUCCESS_COLOR
        )
        self.status_label.config(text=f"✔ Đã sao chép toàn bộ {len(self.current_labels)} nhãn vào Clipboard!")

        def reset_copy_all():
            try:
                self.btn_copy_all.config(
                    text="📋 Copy tất cả",
                    bg="#EFF6FF",
                    fg=PRIMARY_COLOR
                )
            except tk.TclError:
                pass

        self.after(2000, reset_copy_all)

    def save_excel_file(self):
        """Xuất toàn bộ nhãn thành tệp Excel .xlsx chuẩn hóa định dạng."""
        if not self.current_parsed:
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile="nhan_cap_chuan.xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("All files", "*.*")]
        )
        if not file_path:
            return

        try:
            wb = label_export.create_excel_workbook(self.current_parsed)
            if wb:
                wb.save(file_path)
                messagebox.showinfo("Thành công", f"Đã xuất thành công tệp Excel vào:\n{file_path}")
                self.status_label.config(text=f"Đã xuất tệp Excel: {os.path.basename(file_path)}")
            else:
                messagebox.showerror("Lỗi", "Không thể khởi tạo thư viện openpyxl.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể ghi file Excel: {str(e)}")

    def save_to_file(self):
        """Lưu toàn bộ danh sách nhãn ra tệp tin text."""
        if not self.current_labels:
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            initialfile="nhan_cap.txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if not file_path:
            return

        all_content = "\n\n\n".join(b for _, b in self.current_labels)
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(all_content + "\n\n\n")
            messagebox.showinfo("Thành công", f"Đã lưu thành công vào:\n{file_path}")
            self.status_label.config(text=f"Đã lưu danh sách nhãn vào file: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể ghi file: {str(e)}")


if __name__ == "__main__":
    app = CableLabelApp()
    app.mainloop()
