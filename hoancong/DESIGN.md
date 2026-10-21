# Hoàn công CSKVTT — tài liệu thiết kế (v2, 10/2026)

## Mục tiêu đợt này (anh Lâm góp ý 17/10/2026, đã chốt tiêu chí)
| # | Lỗi | Nguyên nhân | Cách sửa | Tiêu chí đạt |
|---|-----|-------------|----------|--------------|
| 1 | Sai ngày cuối tháng ("ngày 31 tháng 9") | Chỉ đổi số tháng, giữ ngày của mẫu | Dòng "ngày DD tháng M năm Y" (không thuộc căn cứ): nếu DD = ngày cuối tháng của mẫu → ngày cuối tháng của kỳ | T9 → "ngày 30 tháng 9"; T2 → 28/29; ngày 01 và ngày trong căn cứ giữ nguyên |
| 2 | Cột "KL kế hoạch Quý" (sheet 3) ghi KL tháng | Code ghi E=F | Đọc **biên bản số 5** (nghiệm thu QLVH tổng hợp, cột "Khối lượng kế hoạch quý") → E sheet 3; nhãn "Quý X/YYYY" theo kỳ | E = số quý trong biên bản; thiếu biên bản → mục kiểm "Xem lại", cột để trống (không lấy số tháng) |
| 3 | Sheet 4 "Vị trí đặt tủ điều khiển" ghi tên đường | Code ghi `duong or diachi` | Ghi `diachi` (cột địa chỉ đặt tủ của biên bản 4) | Khớp cột địa chỉ trong PDF |
| 4 | Số hiệu sheet 8–11 không đổi tháng | Chỉ nhận dạng `01/YCNT-8-Q7`, không nhận `01/08/YCNT` | Mẫu số hiệu STT/Tháng/YCNT (bỏ quận); chỉ đổi phần tháng, giữ cách viết (1 hay 2 chữ số) của mẫu; mẫu cũ `-M-Qn` chuyển sang STT/Tháng/YCNT | Sheet 8–11 và dòng nhắc số phiếu ở 9, 11 đúng tháng kỳ |
| 5 | Chưa có số trang | — | `dinhdang_in`: chân trang phải "Trang &P / &N", nghiêng, Times New Roman 11 (trừ Bìa) | Mọi sheet in có "Trang x / y" |
| 6 | PWA | — | manifest + icon 192/512/maskable/180 + service worker (shell cache-first, pylibs/pyodide cache) + nút cài + hướng dẫn iOS | Lighthouse installable; mở lại khi offline |

Ràng buộc chung: căn cứ pháp lý, Gói thầu, Dự toán, Địa điểm giữ nguyên; 0 lỗi công thức; web (Pyodide) và skill dùng cùng mã `src/`.

## Kiến trúc
`src/refill_lib.py` (đọc PDF, tra tủ) → `src/refill_main.py` (đổ vào mẫu) → `src/phatsinh.py` (BB12/13) → `src/dinhdang_in.py` (in, số trang) ; `src/web_run.py` điều phối trên web; `src/report.py` bảng kiểm.
Nhận diện PDF theo nội dung: 11 cột = BD chi tiết; 9 cột = QLVH chi tiết; 7 cột ≥25 dòng = BoQ 499; **7 cột có "kế hoạch quý" + "Duy trì trạm đèn" = QLVH tổng hợp (biên bản 5, tùy chọn)**; "NGOÀI KẾ HOẠCH" = BB12/13.

## Test
`tests/test_v2.py` — chạy pipeline trên bộ biên bản T08 với kỳ 9 (và kỳ 2), kiểm từng tiêu chí. Chạy:
```
python3 tests/test_v2.py <thư mục PDF> <mẫu .xlsm> <bảng tra .xlsx>
```

## Rủi ro
- Tên hạng mục ở biên bản 5 bị cắt dòng → khớp qua `canon_hm` + khóa không dấu.
- Kỳ tháng 12 → tháng 1 đổi năm: không tự đổi năm (ngoài phạm vi), ghi chú.

## Kiểm trên bộ biên bản THẬT Q10 kỳ 9/2026 (17/10/2026) — phát hiện & sửa
- Số hiệu thật: `01/09/YCNT-Q10` (STT/MM/YCNT-Q<quận>). Anh chốt "đúng như biên bản": có biên bản 8–11 → lấy nguyên số hiệu (`sohieu_pdf`); không có → dựng `STT/MM/YCNT-Q<QUAN>`.
- Bảng hạng mục bảo dưỡng KHÔNG cố định 499: Q8 = 499 (nhóm I 1–95, II 96–499); Q10 = 378 (I 1–55, II 56–375, III "Trung tâm điều khiển" 376–378). Nhóm lấy theo dòng I/II/III của biên bản (`parse_boq` → `grp`); mục kiểm = STT liên tục 1..N.
- BB13 Q10: chữ cột "Vị trí" bị PDF cắt sang ô KL ("u 1,000", "2,000\n)") → `parse13` giữ 1 số, trả phần chữ về đúng dòng của ô Vị trí.
- Một dòng BB13 gốc thiếu ")" ("…(tháo dỡ cáp điện ngầm") — giữ nguyên theo biên bản.
- Test bộ thật: `python3 tests/test_bo_that.py <PDF/> <mẫu> <bảng tra> 10 9` → 12/12; `tests/test_v2.py` (bộ T08) → 17/17.

## Phiên bản & cập nhật (17/10/2026)
- Sự cố: trang mới nhưng trình duyệt/app đã cài vẫn chạy mã xử lý cũ (service worker trả `pylibs.zip` cũ) → báo lỗi "u 1,000" dù đã sửa.
- Sửa: mã xử lý nạp từ `py/pylibs.zip?v=<BUILD>`; service worker lưu theo ĐÚNG phiên bản; trong zip có `build.txt`, worker so khớp với `BUILD` của trang, lệch → báo "bấm Ctrl + F5".
- `VERSION` (vd 2.1.0) + `./build.sh` → `BUILD = VERSION-sha1(src)[:6]`; hiện ở tiêu đề cửa sổ ("Hoàn công CSKVTT v2.1.0"), nhãn cạnh tên app và dòng "Phiên bản …" cuối trang.
- Lỗi số liệu (`phatsinh.num`): ô số lẫn chữ tràn từ cột bên cạnh mà chỉ có 1 số → tự lấy số; còn sai → báo rõ biên bản, file, trang, hạng mục, nguyên nhân, cách kiểm tra, cách khắc phục.

## Đợt góp ý 17/10/2026 (tối) — v2.2.0
- Sheet 10 (và 8) mất mục "3. Thời gian / 4. Địa điểm / câu đề nghị": mã chỉ tìm mục này từ dòng 41 trở xuống; mẫu có bảng ngắn thì mục 3 nằm trên dòng 41 → bị bảng hạng mục ghi đè. Sửa: tìm ngay sau dòng tiêu đề bảng (`find_tail(ws, first)`). Test: mẫu rút ngắn bảng (`mau_ngan`).
- Ô "Phường của quận" tự điền theo mã quận (`src/quan_phuong.py`, NQ 1685/2025; build.sh chép sang index.html). Mã lạ → xoá phần tự điền, giữ phần nhập tay. Python: PREFER trống → lấy theo QUAN.
- Nhãn nhóm lấy ĐÚNG CHỮ biên bản: QLVH "Tủ kết nối về trung tâm"; BD "Công tác duy tu thường xuyên / không thường xuyên", "Trung tâm điều khiển".
- Font một số biên bản (3, 6, 8, 10…) đọc ra mất dấu huyền của "ề" (PDF vẫn hiện đúng): `sua_dau` trả lại cho vê/điêu/nhiêu/đêu (không đụng "đê", "chiêu").
- Bỏ dòng "Phiên bản … Ctrl + F5" cuối trang; số phiên bản giữ ở tiêu đề cửa sổ + nhãn (rê chuột xem bản dựng).
- Test bộ thật: 22/22; T08: 17/17.

## Chuyển báo cáo sự cố (BCSC) sang phường mới — v2.5.0 (19/10/2026)
Tiêu chí đã chốt: đổi theo **tên tủ**; đầu ra **chỉ PDF**; theo góp ý mới nhất **xuất bản sạch để ký lại** (thay cho ghi nối tiếp giữ chữ ký).
- `src/bcsc_font.py` — đọc font TrueType nhúng: BCSC (Crystal Reports) dùng font subset mã riêng, ToUnicode trỏ tới đối tượng hỏng
  → không trích chữ được. Giải mã: mã byte → glyph (cmap 3,0, 0xF000+mã) → **chữ ký glyph** (md5 byte glyf; glyph ghép = chữ gốc + từng dấu)
  → ký tự; chữ có dấu ghép NFC từ chữ gốc + dấu, nên chữ mới gặp lần đầu vẫn đọc được nếu đã biết chữ gốc và dấu.
- `src/bcsc_bang.py` — bảng chữ ký → ký tự, **sinh bởi** `tools/tao_bang_bcsc.py` (chạy offline: so ảnh glyph với Liberation cùng độ rộng;
  cặp "thêm 1 nét" D/Đ, O/Q, c/e, ./, xét riêng nét thêm; dấu so với dấu tổ hợp Liberation sau khi kéo giãn khung; 5 glyph sửa tay `SUA_TAY`).
  Kiểm bằng mắt: `tools/xem_bang_bcsc.py BCSC.pdf <thư mục>` vẽ lưới glyph + chữ đọc được. Gặp BCSC có ký tự lạ (hiện "Cần xem: không đọc được tên tủ")
  → chạy lại `tao_bang_bcsc.py <file mới>` (gộp vào bảng cũ), xem lưới, `./build.sh`.
- `src/bcsc.py` — đọc dòng chữ có toạ độ (pdfplumber + giải mã). `src/bcsc_sua.py` — tìm mục 3, tra tủ, ghi PDF:
  - Tra tủ: khớp không dấu → ưu tiên trùng đúng dấu (Cừ ≠ Cự) → ưu tiên phường thuộc quận cũ (`quan_phuong`) → còn nhiều phường thì không đổi, báo lại.
  - `kieu='sach'` (mặc định, giao diện dùng): pypdf ghi lại toàn bộ file; xoá đúng lệnh Tj/TJ của 2 dòng cũ (theo toạ độ chân chữ),
    viết dòng mới bằng Liberation Serif nhúng (cùng độ rộng Times New Roman, cùng cỡ chữ, cùng chân chữ); bỏ toàn bộ trường/ô chữ ký số (đã ký và ô trống — anh Lâm yêu cầu bỏ ô trống 19/10)
    + hình chữ ký, xoá luôn đối tượng chữ ký khỏi file; bỏ ToUnicode hỏng. Font mới dựng trực tiếp đối tượng (lỗi cũ: mất /W → chữ giãn 1 em).
  - `kieu='noi_tiep'`: ghi nối tiếp (xref stream, /Prev), phủ trắng dòng cũ — giữ chữ ký bản gốc hợp lệ (không dùng trên giao diện).
- **Ký số không làm thay được** (cần token/khoá riêng của người ký) → người ký ký lại trên file sạch.
- Test: `python3 tests/test_bcsc.py <BCSC.pdf> <bảng tra> [pdf khác]` → 40/40 (đọc chữ, tra tủ, bản sạch, nối tiếp, ca xấu).
- 20/10/2026 (v2.5.1): BCSC Q5 T09 (321 trang) có glyph chưa có trong bảng (font subset khác bộ ký tự) → chạy `tao_bang_bcsc.py` gộp thêm
  (177→198 chữ, 34→41 dấu), kiểm lưới bằng mắt: đúng hết; chuyển 109/109. Trang báo rõ "ký tự lạ — gửi file để bổ sung bảng" khi không đọc được.
- 20/10/2026 (v2.5.2): BCSC Q5 T08 báo 2408 ký tự lạ — cùng Times New Roman nhưng **khác hinting** (bản font khác) nên md5 byte glyph không khớp.
  Thêm chữ ký theo HÌNH nét (`TTFNho.hinh`: điểm cuối contour + cờ on-curve + toạ độ, bỏ instructions) và bảng phụ `BANG_HINH`/`DAU_HINH`;
  tra byte trước, không có thì tra hình. Không cần bổ sung bảng: đọc đủ, chuyển 51/51.
- 20/10/2026 (v2.5.3): bản sạch bỏ luôn **hình chữ ký tay** ở các trang khối ký (ảnh nằm giữa dòng chức danh "… PHỤ TRÁCH …" và dòng họ tên);
  ảnh hiện trường giữ nguyên. BCSC Q1: 40 hình, Q5 T08/T09: 20 hình mỗi file.
- 21/10/2026 (v2.5.4): DSBCSC Phú Nhuận T08 — 9 ký tự lạ (F, Ờ, …) → gộp bảng (198→223 chữ, 41→45 dấu). Bắt lỗi máy nhận Ơ thành O
  (cùng độ rộng): thêm quy tắc "đỉnh glyph cao vượt chữ gốc >80 đơn vị → chữ có râu" + `SUA_TAY`. Chuyển 6/6 (P9, P5 → Đức Nhuận; P2 → Cầu Kiệu).
