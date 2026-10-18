#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cable Label Generator Web App (Streamlit)
-----------------------------------------
Giao diện Web chống lệch dòng/lệch cột khi sao chép và dán (Anti-Misalignment Copy & Paste).
Hỗ trợ xuất Plain text, Tab Excel, HTML Table cho Word/Docs, và tải trực tiếp tệp .xlsx.
"""

import html
import io
import streamlit as st

import label_export
from app import SAMPLE_DATA

st.set_page_config(
    page_title="Tạo Nhãn Cáp Mạng • Chống Lệch Format",
    page_icon="🏷️",
    layout="wide"
)

# Tùy chỉnh CSS để hiển thị văn bản đơn cách chuẩn xác và các khối nhãn đẹp mắt
st.markdown("""
<style>
    .label-content {
        font-family: 'Menlo', 'Monaco', 'Consolas', 'Courier New', monospace !important;
        font-size: 13px !important;
        line-height: 1.45 !important;
        white-space: pre !important;
        overflow-x: auto;
        background-color: #f8fafc;
        color: #1e293b;
        padding: 10px 12px;
        border-radius: 6px;
        border: 1px solid #e2e8f0;
        margin: 0;
    }
    .badge-info {
        background-color: #eff6ff;
        color: #1d4ed8;
        padding: 3px 8px;
        border-radius: 4px;
        font-size: 12px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

st.title("🏷️ Công Cụ Tạo Nhãn Cáp Mạng")
st.caption("Chuyển đổi dữ liệu cổng từ Excel thành nhãn 2 chiều • Tối ưu hóa đa nền tảng chống lệch format khi Paste")

# --- THANH BÊN: CẤU HÌNH ĐÍCH ĐẾN (TARGET PLATFORM) ---
with st.sidebar:
    st.header("🎯 Nơi bạn sẽ dán (Paste) dữ liệu?")
    
    target_preset = st.radio(
        "Chọn ứng dụng đích:",
        options=[
            "📄 Phần mềm In nhãn / File Text / Notepad (12 spaces)",
            "📊 Microsoft Excel (Tách cột / ô chuẩn xác)",
            "📝 Zalo / Teams / Slack / Discord (Khối Markdown ```)",
            "📋 Tự do tùy chỉnh khoảng cách"
        ],
        index=0
    )
    
    col_sep = " " * 12
    align_brackets = True
    is_excel_tab = False
    is_markdown = False
    
    if "Phần mềm In nhãn" in target_preset:
        col_sep = " " * 12
        align_brackets = True
        st.info("💡 **In ấn / Text:** Căn lề số nhãn bằng khoảng trắng monospace, khoảng cách giữa 2 đầu cáp là 12 spaces chuẩn.")
    elif "Microsoft Excel" in target_preset:
        col_sep = "\t"
        is_excel_tab = True
        st.success("💡 **Excel:** Khi copy sẽ sử dụng cấu trúc Tab/HTML Table. Khi dán (`Ctrl+V`), Excel sẽ tự động tách thành các cột riêng biệt, **không bao giờ bị xô lệch font**!")
    elif "Zalo / Teams" in target_preset:
        is_markdown = True
        col_sep = " " * 12
        align_brackets = True
        st.info("💡 **Chat:** Tự động bọc trong khối mã ` ``` `, giúp tin nhắn hiển thị font monospace thẳng hàng tuyệt đối.")
    else:
        align_brackets = st.checkbox("Căn thẳng hàng theo cột [số]", value=True)
        col_sep_mode = st.selectbox("Khoảng cách 2 cột:", ["12 spaces", "8 spaces", "Tab '\\t'"])
        col_sep = "\t" if col_sep_mode == "Tab '\\t'" else (" " * int(col_sep_mode.split()[0]))

    st.divider()
    auto_sort = st.checkbox("🔢 Tự động sắp xếp nhãn tăng dần (1 → n)", value=True)

# --- BỐ CỤC 2 CỘT CHÍNH ---
col_left, col_right = st.columns([1, 1], gap="large")

with col_left:
    st.subheader("📥 Dữ liệu nguồn (Input)")
    
    col_b1, col_b2 = st.columns(2)
    if col_b1.button("✨ Nạp dữ liệu mẫu thử"):
        st.session_state["input_val"] = SAMPLE_DATA
    if col_b2.button("🗑️ Xóa trắng"):
        st.session_state["input_val"] = ""

    default_val = st.session_state.get("input_val", "")
    user_input = st.text_area(
        "Dán văn bản (định dạng Fr: ... [số] To: ... [số]):",
        value=default_val,
        height=520,
        placeholder="Fr: S4_U13-14_SRV số 1_Slot 4_Port 1    [1]    To: S4_U41_ODF_port 1->N2_U13_ODF_port 1    [1]..."
    )

with col_right:
    st.subheader("📤 Danh sách nhãn đã tạo")
    
    if user_input.strip():
        parsed_items = label_export.parse_raw_lines(user_input, auto_sort=auto_sort)
        
        if parsed_items:
            total_count = len(parsed_items)
            
            top_c1, top_c2 = st.columns([1, 1])
            with top_c1:
                st.success(f"✔ Đã tạo thành công **{total_count}** bộ nhãn!")
            
            # Chuẩn bị toàn bộ dữ liệu xuất
            all_blocks = []
            for fr_val, idx, to_val in parsed_items:
                if is_excel_tab:
                    blk = label_export.format_excel_tab_block(fr_val, to_val, idx)
                else:
                    blk = label_export.format_plaintext_block(
                        fr_val, to_val, idx,
                        col_separator=col_sep,
                        align_brackets=align_brackets
                    )
                if is_markdown:
                    blk = f"```text\n{blk}\n```"
                all_blocks.append((idx, blk))
            
            all_text = "\n\n\n".join(b for _, b in all_blocks)
            html_table_all = label_export.format_html_table(parsed_items)

            with top_c2:
                # Nút sao chép toàn bộ
                if st.button("📋 Sao chép TẤT CẢ nhãn", key="btn_copy_all_streamlit", use_container_width=True):
                    label_export.copy_multi_format_to_clipboard(all_text, html_table_all)
                    st.toast(f"✔ Đã copy toàn bộ {total_count} nhãn vào Clipboard!", icon="✅")

            # Các tùy chọn tải về trực tiếp
            dl_c1, dl_c2 = st.columns(2)
            with dl_c1:
                # Tải file .txt
                st.download_button(
                    label="💾 Tải tệp tin .txt",
                    data=all_text + "\n\n\n",
                    file_name="nhan_cap.txt",
                    mime="text/plain",
                    use_container_width=True
                )
            with dl_c2:
                # Tải file .xlsx chuyên nghiệp
                wb = label_export.create_excel_workbook(parsed_items)
                if wb:
                    buf = io.BytesIO()
                    wb.save(buf)
                    buf.seek(0)
                    st.download_button(
                        label="📊 Tải tệp tin Excel (.xlsx)",
                        data=buf.getvalue(),
                        file_name="nhan_cap_chuan.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            st.divider()

            # Hiển thị từng nhãn kèm nút copy thông minh
            for (fr_val, idx, to_val), (_, block) in zip(parsed_items, all_blocks):
                with st.container():
                    col_header, col_btn = st.columns([3, 1])
                    
                    with col_header:
                        st.markdown(f"**🏷️ Nhãn [{idx}]**")
                    
                    with col_btn:
                        is_copied = (st.session_state.get("last_copied_id") == idx)
                        btn_label = "✔ Đã copy" if is_copied else "📋 Copy"
                        btn_type = "primary" if is_copied else "secondary"
                        
                        if st.button(btn_label, key=f"cp_btn_{idx}", type=btn_type, use_container_width=True):
                            single_html = label_export.format_html_table([(fr_val, idx, to_val)])
                            label_export.copy_multi_format_to_clipboard(block, single_html)
                            st.session_state["last_copied_id"] = idx
                            st.toast(f"✔ Đã copy Nhãn [{idx}] (Đa định dạng chống lệch)!", icon="📋")
                            st.rerun()

                    # Hiển thị nội dung trong thẻ pre
                    escaped_block = html.escape(block)
                    st.markdown(
                        f"""<pre class="label-content">{escaped_block}</pre>""",
                        unsafe_allow_html=True
                    )
                    st.write("")
        else:
            st.warning("⚠️ Không tìm thấy dòng nào khớp định dạng: Fr: ... [số] To: ... [số]")
    else:
        st.info("👈 Vui lòng dán dữ liệu vào ô bên trái hoặc bấm 'Nạp dữ liệu mẫu thử'.")
