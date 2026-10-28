# Thesis Table Formatter (Word Web Add-in)

Add-in dành cho Microsoft Word (Task Pane) giúp tự động hóa việc căn chỉnh và định dạng bảng biểu theo quy chuẩn khóa luận tốt nghiệp / bài báo khoa học.

---

## 📋 Tính năng & Quy chuẩn định dạng

### 1. Bảng số liệu khóa luận (Data Table)
1. **Kiểu bảng (Table Style):** `Table Grid` (lưới viền chuẩn).
2. **Độ rộng bảng:** `AutoFit Window` (tự động căn vừa khít độ rộng trang văn bản).
3. **Font chữ toàn bảng:** `Times New Roman`, cỡ chữ `11pt`, màu đen (`#000000`).
4. **Khoảng cách đoạn (Paragraph spacing):** 
   - `Space Before`: `0pt`
   - `Space After`: `0pt`
   - `Line Spacing`: `Single` (dòng đơn - 11pt)
5. **Căn lề dọc (Vertical alignment):** Căn giữa ô theo chiều dọc (`Center`) cho toàn bộ các ô trong bảng.
6. **Hàng 1 (Hàng tiêu đề):** Chữ in đậm (`Bold`), căn giữa (`Center`).
7. **Từ hàng 2 trở đi:**
   - **Cột 1 (Cột nhãn/tiêu đề dòng):** Căn trái (`Left`).
   - **Các cột từ cột 2 trở đi (Cột dữ liệu/số liệu):** Căn phải (`Right`).

### 2. Bảng Quốc hiệu - Tiêu ngữ (Header Table) - Tự động nhận diện
- **Đường viền:** Xóa sạch toàn bộ viền lưới (`No border`), bảng hoàn toàn trong suốt.
- **Căn lề ngang:** Cả 2 cột đều **Căn giữa (Center)**.
- **Căn lề dọc:** Căn trên cùng (`Top`) để 2 bên cân đối thẳng hàng từ dòng đầu.
- **Cột trái (Cơ quan/Trường):** Cỡ chữ 12pt, tên trường/cơ quan cấp trên chữ thường, tên viện/khoa/đơn vị cấp dưới in đậm.
- **Cột phải (Quốc hiệu):**
  - `CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM`: In hoa, in đậm, 12pt.
  - `Độc lập – Tự do – Hạnh phúc`: In đậm, 12pt.
  - Dòng địa danh / ngày tháng: In nghiêng (`Italic`), không đậm, 12pt.
- **Độ rộng bảng:** `AutoFit Window`.

---

## 📁 Cấu trúc thư mục dự án

```text
word-thesis-table-formatter/
├── manifest.xml                 # Khai báo cấu hình Add-in cho Office
├── package.json                 # Cấu hình npm & dependencies
├── webpack.config.js            # Cấu hình Webpack & HTTPS Dev Server
├── assets/                      # Icon hiển thị trên Ribbon và Task Pane
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   └── icon-80.png
├── src/
│   └── taskpane/
│       ├── taskpane.html        # Giao diện người dùng (UI)
│       ├── taskpane.css         # Phong cách giao diện Fluent/Clean
│       └── taskpane.js          # Logic định dạng Word qua Office.js API
└── README.md
```

---

## 🚀 Hướng dẫn chạy thử (Sideload) trên Word for Mac

### Cách 1: Chạy tự động với Node.js / Office Tooling (Khuyên dùng)

1. Mở Terminal và di chuyển vào thư mục dự án:
   ```bash
   cd ~/Desktop/word-thesis-table-formatter
   ```

2. Cài đặt các thư viện cần thiết:
   ```bash
   npm install
   ```

3. Khởi chạy server và tự động mở Word (Sideload):
   ```bash
   npm start
   ```
   > **Lưu ý:** Lần đầu tiên chạy, hệ thống sẽ hỏi cấp quyền tin cậy chứng chỉ bảo mật HTTPS cục bộ (localhost SSL certificate), bạn hãy chọn **Agree/Yes** và nhập mật khẩu máy Mac nếu được yêu cầu. Word for Mac sẽ tự động được mở cùng với Add-in.

4. Để dừng server sau khi kiểm thử xong:
   ```bash
   npm stop
   ```

---

### Cách 2: Sideload thủ công qua thư mục `wef` trên macOS

Nếu bạn muốn chạy server độc lập:

1. Chạy dev server:
   ```bash
   npm run dev-server
   ```
   *(Server sẽ chạy tại `https://localhost:3000`)*

2. Tạo thư mục `wef` (nếu chưa có) và copy file `manifest.xml`:
   ```bash
   mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
   cp ~/Desktop/word-thesis-table-formatter/manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
   ```

3. Mở Microsoft Word trên máy Mac:
   - Tạo một văn bản mới hoặc mở file khóa luận của bạn.
   - Chọn menu **Chèn (Insert)** > **Tiện ích bổ sung của tôi (My Add-ins)**.
   - Dưới mục **Tiện ích dành cho nhà phát triển (Developer Add-ins)**, chọn **"Thesis Table Formatter"**.
   - Trên thanh Ribbon sẽ xuất hiện tab **"Khóa Luận"** với nút **"Format Tables"** để mở Task Pane.
