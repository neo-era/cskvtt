# Hoàn công CSKVTT

Trang web lập **Hồ sơ hoàn công duy tu chiếu sáng** (gói G201NA) từ bộ biên bản nghiệm thu PDF.

Mở: https://neo-era.github.io/cskvtt/hoancong/

## Cách dùng
1. Lần đầu: chọn **file mẫu hoàn công** (.xlsm G201) và **bảng tra tủ** (BANG_TRA_TU_CHUAN hoặc Control_Cabinet .xlsx). Trình duyệt tự nhớ cho các lần sau.
2. Chọn chế độ:
   - **Lập hồ sơ mới**: chọn cả bộ biên bản PDF của kỳ; kỳ có phát sinh ngoài kế hoạch thì chọn thêm **biên bản 12** (tổng hợp) và **13** (chi tiết). Chọn tháng, bấm **Lập hồ sơ**.
   - **Bổ sung vào hồ sơ có sẵn**: chọn file hồ sơ .xlsx đã lập, chọn biên bản 12 và 13, bấm **Bổ sung và định dạng**. Không chọn PDF thì trang chỉ chỉnh định dạng in.
3. Xem bảng kiểm tra, bấm **Tải file Excel**.

Trang tự nhận từng biên bản theo nội dung (không theo tên file). Biên bản 12/13 đổ vào 2 sheet phát sinh ngoài kế hoạch
(nhận theo tiêu đề sheet), chi tiết phải khớp tổng hợp từng hạng mục. Bước cuối luôn chỉnh định dạng in: A4, vừa 1 trang ngang,
vùng in tới hết chữ ký, nới dòng chữ dài.

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
- `src/` – mã nguồn Python: refill_lib.py, refill_main.py, report.py (lập hồ sơ), phatsinh.py (biên bản 12/13), dinhdang_in.py (định dạng in), web_run.py (điều phối 1 lần chạy trên trang). Sửa xong thì đóng gói lại vào `pylibs.zip` (`cd src && zip -j ../pylibs.zip *.py`).
