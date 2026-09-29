# Hoàn công CSKVTT

Trang web lập **Hồ sơ hoàn công duy tu chiếu sáng** (gói G201NA) từ bộ biên bản nghiệm thu PDF.

Mở: https://neo-era.github.io/cskvtt/hoancong/

## Cách dùng
1. Lần đầu: chọn **file mẫu hoàn công** (.xlsm G201) và **bảng tra tủ** (Control_Cabinet .xlsx). Trình duyệt tự nhớ cho các lần sau.
2. Chọn cả bộ biên bản PDF của kỳ (trang tự nhận ra file QLVH chi tiết, bảo dưỡng 499 hạng mục, bảo dưỡng chi tiết theo nội dung bảng).
3. Chọn tháng, bấm **Lập hồ sơ** (khoảng 1 phút), xem bảng kiểm tra, bấm **Tải file Excel**.

Mọi xử lý chạy trong trình duyệt (Python/Pyodide). Biên bản, file mẫu, bảng tra **không tải lên đâu cả**
và **không lưu trong repo này** (repo công khai).

## Quy tắc giữ nguyên
- Căn cứ pháp lý, Gói thầu, Dự toán, Địa điểm gói thầu: giữ đúng mẫu.
- Chỉ đổi: Địa bàn (phường thực có trong dữ liệu), kỳ tháng, số hiệu phiếu/biên bản, các dòng Biên bản/Nhật ký/Phiếu YCNT trong mục căn cứ.
- Công thức, định dạng, chữ ký giữ đúng mẫu; Excel tự tính lại khi mở.

## Cấu trúc
- `index.html` – giao diện + Web Worker chạy Python.
- `pylibs.zip` – openpyxl, pdfplumber, pdfminer.six (thuần Python) + mã `src/`.
- `pyodide/` – Pyodide 0.26.4.
- `src/` – mã nguồn Python (refill_lib.py, refill_main.py, report.py). Sửa xong thì đóng gói lại vào `pylibs.zip`.
