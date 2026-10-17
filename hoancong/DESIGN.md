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
