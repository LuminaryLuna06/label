# Cable Label Generator (Hệ thống Sinh Nhãn Cáp Viễn Thông)

Hệ thống tạo nhãn cáp 2 cột 4 dòng chuẩn viễn thông, hỗ trợ tự động căn lề monospace, sắp xếp tự nhiên theo chỉ số `[1] → [n]`, sao chép bảo toàn định dạng vào Excel / Notepad, và xuất file `.xlsx` / `.txt`.

---

## 🌟 Tính Năng Nổi Bật (Key Features)

1. **Tự động sắp xếp nhãn tăng dần (Natural Sorting):**
   - Hỗ trợ dữ liệu đầu vào không theo thứ tự (ví dụ: `[5], [1], [10], [2]`).
   - Sắp xếp tự nhiên theo số nguyên và ký tự (`[1] → [2] → ... → [10]`), không bị lỗi thứ tự chữ số học thông thường.
2. **Sao chép bảo toàn định dạng (`copyWithMonospace`):**
   - Tích hợp chuẩn `navigator.clipboard.write([ClipboardItem])` với 2 payload:
     - `text/html`: Bọc thẻ `<pre style="font-family: 'SF Mono', Consolas, 'Courier New', monospace; font-size: 13px; margin: 0;">` giúp giữ khoảng trắng tuyệt đối khi dán vào Microsoft Word, Outlook, OneNote, rich-text editors.
     - `text/plain`: Sử dụng đúng số lượng khoảng trắng căn lề chuẩn khi dán vào Notepad, TextEdit, terminal.
3. **Xuất Excel & Text trực tiếp:**
   - Xuất file `.xlsx` với 2 cột riêng biệt và độ rộng cột tự động điều chỉnh.
   - Xuất file `.txt` nguyên bản căn lề.
4. **Giao diện Responsive thích ứng đa thiết bị:**
   - Hoàn toàn tương thích từ điện thoại di động (iOS/Android), máy tính bảng (iPad/Tablet) đến màn hình máy tính để bàn độ phân giải cao.
   - Hỗ trợ cuộn ngang thẻ xem trước nhãn mượt mà (`-webkit-overflow-scrolling: touch`), nút chạm ngón tay thân thiện (touch-friendly), và thanh công cụ thích ứng.

---

## 🚀 Các Phương Thức Sử Dụng

### 1. Ứng dụng Web React + TypeScript (`frontend/`)

Được xây dựng với Vite + React 19 + TypeScript + Lucide Icons + SheetJS, chạy 100% phía client (không cần backend server):

```bash
# Di chuyển vào thư mục frontend
cd frontend

# Cài đặt thư viện (nếu chưa cài)
bun install
# hoặc: npm install

# Khởi chạy máy chủ phát triển
bun run dev
# hoặc: npm run dev
```

Truy cập: `http://localhost:5173`

---

### 2. Ứng dụng Desktop Python Tkinter (`app.py`)

Giao diện native desktop có hỗ trợ sắp xếp tự động, sao chép định dạng đa năng (HTML + Plain text trên Windows/macOS/Linux), và xuất file Excel:

```bash
python3 app.py
```

---

### 3. Giao diện Web Streamlit (`web_app.py`)

Giao diện web nhanh với bảng điều khiển trực quan:

```bash
streamlit run web_app.py
```

---

### 4. Trang Web Đơn Lẻ Không Cần Cài Đặt (`gui_web.html`)

Mở trực tiếp tệp [`gui_web.html`](file:///Users/dinhtronganh/Documents/Github/label/gui_web.html) trong bất kỳ trình duyệt nào (Chrome, Safari, Edge, Firefox).

---

## 📋 Cấu Trúc Dự Án

```
├── frontend/                     # React + TypeScript web application
│   ├── src/
│   │   ├── utils/
│   │   │   └── labelGenerator.ts # Thuật toán natural sort, 2-column formatting, copyWithMonospace, Excel export
│   │   ├── App.tsx               # Giao diện chính với hiệu ứng Toast & feedback copy
│   │   ├── App.css               # Phong cách tối giản hiện đại (Apple/Slate theme)
│   │   └── types.ts              # Định nghĩa kiểu dữ liệu TypeScript
│   └── package.json
├── app.py                        # Python Tkinter GUI
├── label_export.py               # Core parsing & clipboard utility (Python)
├── web_app.py                    # Streamlit Web App
├── gui_web.html                  # Standalone Single-file HTML/JS App
└── README.md
```
