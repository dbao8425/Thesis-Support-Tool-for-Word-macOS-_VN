# Hướng dẫn cài đặt & sử dụng Add-in trên Windows

Add-in **Duy Bảo - Word Tools** (Thesis Table Formatter) được xây dựng trên nền tảng chuẩn **Microsoft Office.js**, hoạt động mượt mà trên cả **Windows** (Word 2013, 2016, 2019, 2021, Office 365) và **macOS**.

Toàn bộ giao diện và logic đã được đưa lên máy chủ đám mây GitHub Pages (HTTPS bảo mật cao), vì vậy bạn **không cần cài đặt Node.js hay bật bất kỳ server nào** trên máy tính Windows.

---

## ⚡ Cách 1: Cài đặt tự động trong 1 giây (Khuyên dùng)

### Bước 1: Tải bộ cài về máy
Tải file `Cai_Dat_Windows.bat` (hoặc tải toàn bộ thư mục/file zip này) về máy tính Windows của bạn.

### Bước 2: Chạy cài đặt
1. **Nhấp đúp chuột vào file `Cai_Dat_Windows.bat`**.
2. Màn hình cài đặt màu đen sẽ tự động:
   - Tải/sao chép tệp cấu hình `manifest.xml` vào máy.
   - Tự động đăng ký Add-in vào Word cho Windows.
   - Làm mới bộ nhớ đệm (cache).
3. Khi thấy thông báo **"🎉 CÀI ĐẶT THÀNH CÔNG CHO WINDOWS! 🎉"**, bạn bấm phím bất kỳ để đóng cửa sổ.

### Bước 3: Mở Word và sử dụng
1. Mở phần mềm **Microsoft Word** trên Windows (nếu Word đang mở trước đó, hãy đóng lại rồi mở lại).
2. Nhìn lên thanh công cụ Ribbon trên cùng, bạn sẽ thấy tab mới tên **"Duy Bảo"**.
3. Bấm vào nút **"Format Tables"** để mở bảng điều khiển bên phải.

---

## 🛠️ Cách 2: Cài đặt qua Thư mục chia sẻ mạng (Shared Folder)
*(Dành cho máy tính cơ quan, trường học bị khóa quyền ghi Registry)*

1. Tạo một thư mục trên ổ đĩa `C:\`, ví dụ: `C:\WordAddins`.
2. Sao chép tệp `manifest.xml` vào thư mục `C:\WordAddins`.
3. Nhấp chuột phải vào thư mục `WordAddins` > chọn **Properties** > chọn thẻ **Sharing** > bấm **Share...**.
4. Chọn tài khoản Windows của bạn (hoặc `Everyone`) > bấm **Share** > bấm **Done**.
5. Ghi nhớ đường dẫn mạng (dạng `\\<Tên-Máy-Tính>\WordAddins` hoặc `\\localhost\WordAddins`).
6. Mở **Microsoft Word** > vào menu **File (Tệp)** > **Options (Tùy chọn)**.
7. Chọn **Trust Center (Trung tâm tin cậy)** > **Trust Center Settings (Cài đặt trung tâm tin cậy)**.
8. Chọn thẻ **Trusted Add-in Catalogs (Danh mục phần bổ trợ tin cậy)**:
   - Nhập đường dẫn `\\localhost\WordAddins` vào ô **Catalog Url**.
   - Bấm nút **Add Catalog (Thêm danh mục)**.
   - Đánh dấu tick vào ô **Show in Menu (Hiện trong menu)** bên cạnh.
   - Bấm **OK** và khởi động lại Word.
9. Trong Word, vào tab **Insert (Chèn)** > **Add-ins (Phần bổ trợ)** > **My Add-ins (Phần bổ trợ của tôi)** > chọn tab **SHARED FOLDER (Thư mục dùng chung)** > chọn **Duy Bảo - Word Tools** > bấm **Add**.

---

## 🌐 Cách 3: Sử dụng trên Word Online (Web)
*(Dành cho người dùng Microsoft 365 muốn dùng ngay trên trình duyệt mà không cần cài vào máy)*

1. Truy cập [https://word.office.com](https://word.office.com) và mở tài liệu của bạn.
2. Vào thẻ **Insert (Chèn)** > **Add-ins (Phần bổ trợ)** > chọn **Upload My Add-in (Tải lên phần bổ trợ của tôi)**.
3. Chọn tệp `manifest.xml`.
4. Tab **Duy Bảo** sẽ xuất hiện ngay lập tức trên thanh Ribbon trình duyệt.

---

## 🗑️ Cách gỡ cài đặt (Uninstall)

Nếu không còn nhu cầu sử dụng, bạn chỉ cần:
- Nhấp đúp chuột vào file `Go_Cai_Dat_Windows.bat`.
- Hệ thống sẽ tự động gỡ bỏ sạch sẽ Add-in khỏi Word.
