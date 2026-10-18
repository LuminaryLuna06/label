# Cable Label Generator (Hệ thống Sinh Nhãn Cáp Viễn Thông)

Hệ thống tạo nhãn cáp 2 cột 4 dòng chuẩn viễn thông, hỗ trợ tự động căn lề monospace, sắp xếp tự nhiên theo chỉ số `[1] → [n]`, sao chép bảo toàn định dạng vào Excel / Notepad, và xuất file `.xlsx` / `.txt`.

---

## 🌟 Tính Năng Nổi Bật (Key Features)

1. **Lưu trữ dữ liệu nguồn tự động (LocalStorage):**
   - Dữ liệu nhập được lưu trữ tức thì vào LocalStorage trình duyệt.
   - Khi tải lại hoặc mở lại trang web, toàn bộ dữ liệu nguồn và tùy chọn cấu hình được khôi phục nguyên vẹn thay vì phải sử dụng dữ liệu mẫu mặc định.
2. **Tự động sắp xếp nhãn tăng dần (Natural Sorting):**
   - Hỗ trợ dữ liệu đầu vào không theo thứ tự (ví dụ: `[5], [1], [10], [2]`).
   - Sắp xếp tự nhiên theo số nguyên và ký tự (`[1] → [2] → ... → [10]`), không bị lỗi thứ tự chữ số thông thường.
3. **Sao chép bảo toàn định dạng (`copyWithMonospace`):**
   - Tích hợp chuẩn `navigator.clipboard.write([ClipboardItem])` với 2 payload:
     - `text/html`: Bọc thẻ `<pre style="font-family: 'SF Mono', Consolas, 'Courier New', monospace; font-size: 13px; margin: 0;">` giúp giữ khoảng trắng tuyệt đối khi dán vào Microsoft Word, Outlook, OneNote, rich-text editors.
     - `text/plain`: Sử dụng đúng số lượng khoảng trắng căn lề chuẩn khi dán vào Notepad, TextEdit, terminal.
4. **Xuất Excel & Text trực tiếp:**
   - Xuất file `.xlsx` với 2 cột riêng biệt và độ rộng cột tự động điều chỉnh.
   - Xuất file `.txt` nguyên bản căn lề.
5. **Giao diện Responsive thích ứng đa thiết bị:**
   - Hoàn toàn tương thích từ điện thoại di động (iOS/Android), máy tính bảng (iPad/Tablet) đến màn hình máy tính để bàn độ phân giải cao.
   - Hỗ trợ cuộn ngang thẻ xem trước nhãn mượt mà (`-webkit-overflow-scrolling: touch`), nút chạm ngón tay thân thiện (touch-friendly), và thanh công cụ thích ứng.

---

## 🚀 Cài Đặt & Vận Hành

Ứng dụng web được xây dựng với Vite + React 19 + TypeScript + Lucide Icons + SheetJS, chạy 100% phía client (0ms latency, không phụ thuộc backend nặng nề):

```bash
# Di chuyển vào thư mục frontend
cd frontend

# Cài đặt thư viện
npm install
# hoặc: bun install

# Khởi chạy môi trường phát triển
npm run dev

# Hoặc build và chạy bản preview sản xuất
npm run build
npm run preview -- --host 0.0.0.0 --port 5175
```

---

## 📋 Cấu Trúc Dự Án Tối Giản (Lightweight Architecture)

```
├── frontend/                     # React + TypeScript web application
│   ├── src/
│   │   ├── utils/
│   │   │   └── labelGenerator.ts # Thuật toán natural sort, 2-column formatting, copyWithMonospace, Excel export
│   │   ├── App.tsx               # Giao diện chính với LocalStorage, Toast & feedback copy
│   │   ├── App.css               # Phong cách tối giản hiện đại (Apple/Slate theme) & responsive
│   │   ├── types.ts              # Định nghĩa kiểu dữ liệu TypeScript
│   │   └── main.tsx              # Entry point
│   ├── index.html
│   ├── vite.config.ts            # Cấu hình Vite & allowedHosts cho reverse proxy
│   └── package.json
├── .gitignore                    # Loại bỏ cache và build artifacts
└── README.md
```
