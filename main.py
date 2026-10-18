import re


def generate_cable_labels(raw_text: str, output_file: str = "nhan_cap.txt"):
    # Chuẩn hóa khoảng trắng đặc biệt (non-breaking space \xa0) thành space thường
    cleaned_text = raw_text.replace("\xa0", " ")
    lines = cleaned_text.strip().split("\n")

    blocks = []

    # Regex bắt định dạng: Fr: <nội dung> [số] To: <nội dung> [số]
    pattern = re.compile(
        r"Fr:\s*(.+?)\s*\[(\d+)\]\s+To:\s*(.+?)\s*\[(\d+)\]", re.IGNORECASE
    )

    for line in lines:
        line = line.strip()
        if not line:
            continue

        match = pattern.search(line)
        if match:
            # Lấy thông tin gốc và dọn sạch khoảng trắng thừa ở 2 đầu
            fr_val = match.group(1).strip()
            idx = match.group(2).strip()
            to_val = match.group(3).strip()

            # Tạo chuỗi cơ sở cho 2 chiều
            # Chiều thuận (Đầu 1)
            f1 = f"Fr: {fr_val}"
            t1 = f"To: {to_val}"

            # Chiều nghịch (Đầu 2 - Đảo Fr và To)
            f2 = f"Fr: {to_val}"
            t2 = f"To: {fr_val}"

            # Tìm độ dài lớn nhất trong 4 dòng của cùng 1 nhãn để làm mốc căn lề
            max_len = max(len(f1), len(t1), len(f2), len(t2))

            # Hàm con tạo từng dòng hoàn chỉnh (căn lề + [số] + 12 spaces + lặp lại cột 2)
            def format_row(left_text: str, label_id: str) -> str:
                # Dòng dài nhất sẽ có (max_len - len) = 0, cộng thêm 2 khoảng trắng mặc định
                padding = " " * (max_len - len(left_text) + 2)
                single_col = f"{left_text}{padding}[{label_id}]"
                return f"{single_col}{' ' * 12}{single_col}"

            # Ghép thành khối 4 dòng cho 1 sợi cáp
            block = (
                f"{format_row(f1, idx)}\n"
                f"{format_row(t1, idx)}\n\n"
                f"{format_row(f2, idx)}\n"
                f"{format_row(t2, idx)}"
            )
            blocks.append((idx, block))

    # Xuất kết quả ra màn hình và ghi vào file .txt
    with open(output_file, "w", encoding="utf-8") as f:
        for idx, b in blocks:
            # In ra console có kèm tiêu đề để dễ nhìn
            print(f"--- Nhãn [{idx}] ---")
            print(b)
            print()

            # Ghi vào file text (chỉ ghi nội dung nhãn, cách nhau 2 dòng trống)
            f.write(b + "\n\n\n")

    print(f"=> Đã xuất thành công {len(blocks)} bộ nhãn vào file '{output_file}'!")


# =====================================================================
# DÁN DỮ LIỆU COPY TỪ EXCEL VÀO GIỮA CẶP DẤU ''' BÊN DƯỚI
# =====================================================================
raw_data = """
Fr: S4_U13-14_SRV số 1_Slot 4_Port 1    [1]	To: S4_U41_ODF_port 1->N2_U13_ODF_port 1    [1]
Fr: S4_U13-14_SRV số 1_Slot 4_Port 2    [2]	To: S4_U41_ODF_port 2->N2_U13_ODF_port 2    [2]
Fr: S4_U13-14_SRV số 1_iDrac    [3]	To: S4_U40_PP_Port 5->N1_U3_PP_Port 5    [3]
Fr: N2_U13_ODF_port 1    [4]	To: N3_U1-13_CoreSw01_Slot 10_port 1    [4]
Fr: N2_U13_ODF_port 2    [5]	To: N4_U1-13_CoreSw02_Slot 10_port 1    [5]
Fr: N1_U3_PP_Port 5    [6]	To: N1_U20_SWMgmt_port 18    [6]
	
Fr: S4_U11-12_SRV số 2_Slot 4_Port 1    [7]	To: S4_U41_ODF_port 3->N2_U13_ODF_port 3    [7]
Fr: S4_U11-12_SRV số 2_Slot 4_Port 2    [8]	To: S4_U41_ODF_port 4->N2_U13_ODF_port 4    [8]
Fr: S4_U11-12_SRV số 2_iDrac    [9]	To: S4_U40_PP_Port 8->N1_U3_PP_Port 8    [9]
Fr: N2_U13_ODF_port 3    [10]	To: N3_U1-13_CoreSw01_Slot 10_port 2    [10]
Fr: N2_U13_ODF_port 4    [11]	To: N4_U1-13_CoreSw02_Slot 10_port 2    [11]
Fr: N1_U3_PP_Port 8    [12]	To: N1_U20_SWMgmt_Port 19    [12]
	
Fr: S4_U9-10_SRV số 3_Slot 4_Port 1    [13]	To: S4_U41_ODF_port 5->N2_U13_ODF_port 5    [13]
Fr: S4_U9-10_SRV số 3_Slot 4_Port 2    [14]	To: S4_U41_ODF_port 6->N2_U13_ODF_port 6    [14]
Fr: S4_U9-10_SRV số 3_iDrac    [15]	To: S4_U40_PP_Port 9->N1_U3_PP_Port 9    [15]
Fr: N2_U13_ODF_port 5    [16]	To: N3_U1-13_CoreSw01_Slot 10_port 3    [16]
Fr: N2_U13_ODF_port 6    [17]	To: N4_U1-13_CoreSw02_Slot 10_port 3    [17]
Fr: N1_U3_PP_Port 9    [18]	To: N1_U20_SWMgmt_Port 20    [18]
	
Fr: S4_U25_SW-7710B số 1_Port 1    [19]	To: S4_U2-3_SAN 5224_A0    [19]
Fr: S4_U25_SW-7710B số 1_Port 2    [20]	To: S4_U2-3_SAN 5224_A2    [20]
Fr: S4_U25_SW-7710B số 1_Port 3    [21]	To: S4_U2-3_SAN 5224_B0    [21]
Fr: S4_U25_SW-7710B số 1_Port 4    [22]	To: S4_U2-3_SAN 5224_B2    [22]
Fr: S4_U25_SW-7710B số 1_Port 23    [23]	To: S4_U41_ODF_Port 7->N2_U13_ODF_Port 7    [23]
Fr: S4_U25_SW-7710B số 1_Port 24    [24]	To: S4_U41_ODF_Port 8->N2_U13_ODF_Port 8    [24]
Fr: S4_U25_SW-7710B số 1_Mgmt    [25]	To: S4_U40_PP_0_Port 10    [25]
Fr: S4_U25_SW-7710B số 1_Mgmt    [26]	To: S4_U40_PP_0_Port 10    [26]
Fr: N2_U13_ODF_Port 7    [27]	To: N3_U1-13_CoreSw 01_Slot 10_Port 23    [27]
Fr: N2_U13_ODF_Port 8    [28]	To: N3_U1-13_CoreSw 01_Slot 10_Port 23    [28]
	
Fr: S4_U23_SW-7710B số 2_Port 1    [29]	To: S4_U2-3_SAN 5224_A1    [29]
Fr: S4_U23_SW-7710B số 2_Port 2    [30]	To: S4_U2-3_SAN 5224_A3    [30]
Fr: S4_U23_SW-7710B số 2_Port 3    [31]	To: S4_U2-3_SAN 5224_B1    [31]
Fr: S4_U23_SW-7710B số 2_Port 4    [32]	To: S4_U2-3_SAN 5224_B3    [32]
Fr: S4_U23_SW-7710B số 2_Port 23    [33]	To: S4_U41_ODF_Port 9->N2_U13_ODF_Port 9    [33]
Fr: S4_U23_SW-7710B số 2_Port 24    [34]	To: S4_U41_ODF_Port 10->N2_U13_ODF_Port 10    [34]
Fr: S4_U23_SW-7710B số 2_Mgmt    [35]	To: S4_U40_PP_0_Port 11    [35]
Fr: S4_U23_SW-7710B số 2_Mgmt    [36]	To: S4_U40_PP_0_Port 11    [36]
Fr: N2_U13_ODF_Port 9    [37]	To: N3_U1-13_CoreSw 01_Slot 10_Port 24    [37]
Fr: N2_U13_ODF_Port 10    [38]	To: N3_U1-13_CoreSw 01_Slot 10_Port 24    [38]
	
Fr: S4_U2-3_SAN 5224_Ctrl-A_Mgmt    [39]	To: S4_U40_PP_Port 12->N1_U3_PP_Port 12    [39]
Fr: S4_U2-3_SAN 5224_Ctrl-B_Mgmt    [40]	To: S4_U40_PP_Port 13->N1_U3_PP_Port 13    [40]
Fr: N1_U3_PP_Port 12    [41]	To: N1_U20_Sw-Mgmt_Port 23    [41]
Fr: N1_U3_PP_Port 13    [42]	To: N1_U20_Sw-Mgmt_Port 24    [42]
"""

if __name__ == "__main__":
    generate_cable_labels(raw_data)
