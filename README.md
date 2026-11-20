# Duy Bảo - Word Tools (Thesis Table Formatter)

Add-in dành cho Microsoft Word (Task Pane) giúp tự động hóa việc căn chỉnh và định dạng bảng biểu theo quy chuẩn khóa luận tốt nghiệp / bài báo khoa học Việt Nam.

Add-in hoạt động hoàn toàn trên nền tảng đám mây (GitHub Pages), tương thích 100% với cả **Windows** và **macOS** mà không cần cài đặt Node.js hay bật server cục bộ.

---

## 📋 Tính năng nổi bật & Quy chuẩn định dạng

### 1. Bảng số liệu khóa luận (Data Table)
1. **Kiểu bảng (Table Style):** `Table Grid` (lưới viền đen đơn chuẩn mực).
2. **Font chữ & Cỡ chữ:** `Times New Roman`, tùy chỉnh cỡ chữ linh hoạt từ `6pt` đến `16pt` (mặc định `11pt`, có nút tăng giảm `±0.5pt` và nút chọn nhanh `9pt`, `10pt`, `11pt`, `12pt`).
3. **Lề ô (Cell Margin) tùy chỉnh 4 cạnh độc lập:**
   - Hỗ trợ nhập riêng từng cạnh: **Trên**, **Dưới**, **Trái**, **Phải** (đơn vị `cm`).
   - Có nút tăng giảm `±0.05cm` và gợi ý nhanh: `Đều 0.1cm`, `Đều 0.15cm`, `Đều 0.2cm (Chuẩn)`, `T/D 0.15 T/P 0.1`.
4. **Kích thước bảng (Table Dimensions - Tùy chọn):**
   - **Độ cao hàng (Row Height):** Nhập số đo theo `cm` (tăng giảm `±0.1cm`).
   - **Độ rộng bảng (Table Width):** Nhập số đo theo `cm` (tăng giảm `±0.5cm`).
   - Có nút **"Bỏ trống"** để giữ nguyên kích thước ban đầu của bảng nếu không muốn can thiệp.
5. **Tự động căn đều kích thước các cột số liệu:**
   - Chia đều chiều rộng các cột dữ liệu (từ Cột 2 trở đi) theo độ rộng tùy chỉnh (hoặc khổ A4 chuẩn) nhằm đảm bảo số liệu không bị tràn hay nhảy dòng.
   - Tự động điều chỉnh tỷ lệ cột nhãn (Cột 1) phù hợp theo số lượng cột.
6. **Khoảng cách đoạn & Lùi đầu dòng (Paragraph formatting):**
   - `Space Before`: `0pt`
   - `Space After`: `0pt`
   - `Line Spacing`: `Single` (dòng đơn theo cỡ chữ)
   - `First Line Indent`: `0pt` (loại bỏ hoàn toàn lùi đầu dòng cho toàn bộ văn bản trong bảng)
   - `Left Indent` / `Right Indent`: `0pt` (loại bỏ thụt lề)
6. **Căn lề:**
   - **Hàng 1 (Hàng tiêu đề):** In đậm (`Bold`), Căn giữa (`Center`).
   - **Cột 1 (Từ hàng 2 trở đi):** Căn trái (`Left`).
   - **Cột 2 trở đi (Từ hàng 2 trở đi):** Căn phải (`Right`).
   - **Toàn bộ ô trong bảng:** Căn giữa ô theo chiều dọc (`Vertical Center`).

### 2. Bảng Quốc hiệu - Tiêu ngữ (Header Table) - Tự động nhận diện
- **Đường viền:** Xóa sạch toàn bộ viền lưới (`No border`), bảng hoàn toàn trong suốt.
- **Căn lề ngang:** Cả 2 cột đều **Căn giữa (Center)**.
- **Căn lề dọc:** Căn trên cùng (`Top`) để 2 bên cân đối thẳng hàng từ dòng đầu.
- **Cột trái (Cơ quan/Trường):** Cỡ chữ 12pt, tên trường/cơ quan cấp trên chữ thường, tên viện/khoa/đơn vị cấp dưới in đậm.
- **Cột phải (Quốc hiệu):**
  - `CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM`: In hoa, in đậm, 12pt.
  - `Độc lập – Tự do – Hạnh phúc`: In đậm, 12pt.
  - Dòng địa danh / ngày tháng: In nghiêng (`Italic`), không đậm, 12pt.

---

## 🪟 Hướng dẫn cài đặt trên WINDOWS (1 Click)

### Cách 1: Tự động (Khuyên dùng)
1. Tải về file `Cai_Dat_Windows.bat` (hoặc tải trọn bộ source code).
2. **Nhấp đúp chuột vào `Cai_Dat_Windows.bat`**.
3. Cửa sổ hiện **"🎉 CÀI ĐẶT THÀNH CÔNG CHO WINDOWS! 🎉"** -> Bấm phím bất kỳ để hoàn tất.
4. Mở Word trên máy tính -> Thấy tab **"Duy Bảo"** trên thanh Ribbon -> Bấm **"Format Tables"**.

*(Chi tiết xem thêm tại [HUONG_DAN_WINDOWS.md](HUONG_DAN_WINDOWS.md))*

---

## 🍎 Hướng dẫn cài đặt trên macOS

1. Mở Terminal và copy file `manifest.xml` vào thư mục add-in của Word:
   ```bash
   mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
   cp manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
   ```
2. Mở Microsoft Word trên máy Mac.
3. Tab **"Duy Bảo"** sẽ xuất hiện trên thanh Ribbon với nút **"Format Tables"**.
*(Nếu không thấy, vào menu Chèn/Insert > My Add-ins > Developer Add-ins > chọn Duy Bảo - Word Tools)*.

---

## 🌐 Sử dụng trên Word Online (Trình duyệt)

Dành cho người dùng Microsoft 365:
1. Mở tài liệu trên [https://word.office.com](https://word.office.com).
2. Vào tab **Insert** > **Add-ins** > **Upload My Add-in** > chọn `manifest.xml`.
3. Tab **Duy Bảo** sẽ xuất hiện ngay lập tức trên thanh công cụ.

---

## 📁 Cấu trúc dự án

```text
Thesis-Support-Tool-for-Word-macOS-_VN/
├── manifest.xml                 # Khai báo cấu hình Add-in (trỏ về GitHub Pages)
├── Cai_Dat_Windows.bat          # File cài đặt tự động 1-click cho Windows
├── Go_Cai_Dat_Windows.bat       # File gỡ cài đặt sạch sẽ cho Windows
├── Cai_Dat_Windows.ps1          # Script cài đặt qua PowerShell
├── HUONG_DAN_WINDOWS.md         # Hướng dẫn chi tiết cho Windows
├── assets/                      # Icon hiển thị trên Ribbon và Task Pane
├── dist/                        # Bản build tĩnh triển khai trên gh-pages
├── src/taskpane/                # Mã nguồn HTML, CSS, JavaScript (Office.js)
└── README.md
```
