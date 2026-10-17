# Nút "Tải hồ sơ PDF" — thiết kế & kiểm chứng (v2.3.0, 17/10/2026)

## Tiêu chí (anh Lâm duyệt — Chốt 1)
1. Bấm 1 nút → 1 file PDF gồm mọi sheet đang hiện (bỏ sheet/dòng/cột ẩn), đúng thứ tự.
2. Trang in như Excel: A4 ngang/dọc, lề, vùng in, vừa 1 trang ngang, lặp dòng tiêu đề, chân trang "Trang x / y" (trừ Bìa).
3. Số tính đúng (ô =SUM…), hiện kiểu Việt Nam (5,000 / 1.234,50).
4. Tiếng Việt đủ dấu, chọn/tìm được chữ; font cùng khổ Times New Roman (Liberation Serif) → ngắt dòng, ngắt trang như Excel.
5. Đối chiếu tự động với LibreOffice.
6. Ca xấu: ô quá dài, công thức lạ (báo rõ ô), mất mạng sau lần đầu.
7. Bộ Q10 (~1.900 dòng) < 30 giây.

## Kiến trúc
- `src/xuat_pdf.py` (thuần Python, chạy cả Pyodide và máy): `dinh_dang` (định dạng số VN) · `TinhCongThuc` (SUM/MIN/MAX/ROUND/IF, + − × ÷, tham chiếu sheet khác; hàm lạ → để trống + cảnh báo `Sheet!Ô`) · `TTF` (đọc cmap/hmtx, cắt gọn font chỉ giữ chữ dùng) · `PDF` (Type0/CIDFontType2 + ToUnicode → chọn/tìm chữ được) · `Sheet`/`ve_sheet` (dàn trang kiểu Excel).
- `src/do_rong_chu.py` (sinh bởi `tools/tao_do_rong_chu.py`): bề rộng chữ của font thay thế — để **đo/ngắt dòng không cần tải font** (dùng cho định dạng in Excel).
- `fonts/`: Liberation Serif (Times New Roman), Carlito (Calibri) — SIL OFL, xem `fonts/LICENSE.txt`. Service worker lưu cache-first (`hc-pdffonts`), cài sẵn 4 font serif khi cài app.
- Trang: worker nhận `{type:'pdf'}` → tải 4 font serif song song (lần đầu) → `xuat_pdf.xuat('/work/run/out.xlsx', …)`; thiếu font khác (Carlito) → tải rồi chạy lại.

## Mô hình in (khớp Excel)
- Cột: px = trunc(((256·w + trunc(128/MDW))/256)·MDW), MDW = bề rộng chữ số lớn nhất của font Normal (Calibri 11 → 7px); pt = px·0,75.
- Dòng: chiều cao đặt sẵn; dòng không đặt → tự canh như Excel. Dòng/trang: TNR 12 → 15,75pt (×1,3125), Calibri → ×1,3636.
- Lề trong ô 1,5pt. **Hiệu chỉnh bằng chính Excel**: 1.809 dòng do Excel tự canh chiều cao trong file mẫu → ngắt dòng khớp **99,4%** (lệch còn lại cân bằng 5 nhiều/6 ít).
- Co giãn vừa trang: floor(%) như Excel; không in trang trống; viền bên trong ô gộp không vẽ; chữ tràn sang ô trống bên cạnh; số không vừa → ####.

## Phát hiện & sửa kèm theo (ảnh hưởng cả file Excel)
- Chiều cao dòng do `dinhdang_in` nới bằng hệ số K=0,42 → thấp hơn Excel cần ở ~60 ô (sheet 7: 26 ô) → **chữ bị cắt khi in Excel**. Nay nới bằng cùng mô hình đo chữ ở trên (kể cả Bìa).
- Vùng in mẫu dạng chỉ cột (`'4_…'!$C:$K`) openpyxl không đọc → sheet 4 in cả cột phụ AC (bảng bị thu nhỏ). Nay đọc thẳng từ file mẫu (`_vung_in_mau`).
- Nhãn Gói thầu/Dự toán/Địa điểm đo chiều cao SAU khi thành đậm + gạch dưới.
- Chữ dấu rời (tổ hợp) trong mẫu → chuẩn hoá NFC khi vẽ (dấu không lệch).

## Kiểm chứng (`tests/test_pdf.py <hồ sơ.xlsx> <mẫu.xlsm>`)
26/26 đạt trên hồ sơ Q10 T09 đầy đủ (có BB12/13): định dạng số; 1.050 ô công thức khớp LibreOffice; chữ của 9.500+ ô nằm trọn trong khung ô trên trang (ô bị cắt đều có cảnh báo); chân trang theo từng biên bản; khổ giấy; tiếng Việt; không còn ô bị cắt chữ do chiều cao; ngắt dòng khớp Excel 99,4%; ca xấu (VLOOKUP → báo `10.PYCNT2_BD_G201!B34`, ô 6.000 ký tự, không vùng in, dòng ẩn).
- Số trang so với LibreOffice: chênh ≤ 2/sheet (6: 18/17, 7: 69/67, 10: 4/3). Nguyên nhân phía LibreOffice: co 98,4% thay vì 100% như Excel; ở sheet 6 LibreOffice còn **in tràn** (23 hạng mục nằm ngoài khổ giấy).
- Trình duyệt (Chromium): 152 trang, ~11–13 giây xử lý (lần đầu ~25 giây gồm tải font); tắt mạng tải lại → vẫn lập + xuất PDF được.

## Còn lưu ý
- Sheet 9, 11 (mẫu): vài dòng chữ dài hơn vùng in A:F ("+ Bên A: TRUNG TÂM … TP. HỒ CHÍ MINH", "Chức vụ: … Chiếu sáng số 1") → bị cắt khi in, cả trên Excel. Sửa ở mẫu (nới cột/thu chữ) — PDF liệt kê trong "lưu ý khi in".
- Font thay thế: Liberation Serif cùng số đo với Times New Roman nhưng nét chữ hơi khác.
