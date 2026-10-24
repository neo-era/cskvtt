# Gắn Google Sheet nhật ký sử dụng (làm 1 lần, ~5 phút)

1. Tạo 1 Google Sheet mới (vd "Nhật ký app Hoàn công CSKVTT") bằng tài khoản Google của anh.
2. Trong Sheet: **Tiện ích mở rộng → Apps Script**. Xoá hết mã mẫu, dán toàn bộ nội dung file `Code.gs` này, bấm **Lưu**.
3. Bấm **Triển khai → Tùy chọn triển khai mới** → loại **Ứng dụng web**:
   - *Thực thi với tư cách*: **Tôi**
   - *Người có quyền truy cập*: **Bất kỳ ai**
   Bấm **Triển khai**, cấp quyền khi Google hỏi (chọn tài khoản → Nâng cao → Đi tới … → Cho phép).
4. Sao chép **URL ứng dụng web** (dạng `https://script.google.com/macros/s/…/exec`) gửi cho Claude để gắn vào app.

Sheet tự tạo trang tính **NhatKy** với dòng tiêu đề ở lần ghi đầu.
- Script chỉ **ghi thêm** dòng, không đọc/sửa/xoá được dữ liệu cũ.
- Người dùng tự khai tên (không xác thực).
- Đường dẫn nằm trong mã nguồn công khai: ai có đường dẫn cũng ghi thêm dòng được, nhưng không đọc được.
