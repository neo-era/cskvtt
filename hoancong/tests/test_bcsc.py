# -*- coding: utf-8 -*-
"""Kiem tinh nang chuyen bao cao su co (BCSC) sang phuong moi.
Chay: python3 tests/test_bcsc.py <BCSC.pdf> <bang tra .xlsx|.json> [pdf_khac_khong_phai_bcsc.pdf]"""
import os, sys, subprocess, tempfile, shutil
HERE = os.path.dirname(os.path.abspath(__file__)); SRC = os.path.join(HERE, '..', 'src'); FONTS = os.path.join(HERE, '..', 'fonts')
sys.path.insert(0, SRC)
from bcsc import DocBCSC
from bcsc_sua import chuyen, tim_bao_cao, TraTu, doc_bang_tra

PDF, BT = sys.argv[1], sys.argv[2]
KHAC = sys.argv[3] if len(sys.argv) > 3 else None
goc = open(PDF, 'rb').read()
rows = doc_bang_tra(BT)
ok = 0; fail = 0
def T(ten, dk, chi=''):
    global ok, fail
    if dk: ok += 1; print('  ĐẠT  ', ten)
    else: fail += 1; print('  LỖI  ', ten, chi)

print('1. Đọc chữ (font mã hoá riêng)')
tl = DocBCSC(goc); bcs = tim_bao_cao(tl)
T('tìm đủ 78 báo cáo', len(bcs) == 78, len(bcs))
b0 = bcs[0]
T('báo cáo 1: mã, tủ, phường, quận', (b0['ma'], b0['tu'], b0['phuong']['text'], b0['quan']['text']) ==
  ('353270', 'Công Viên 23/9 (Khu B) -1', '- Phường Phạm Ngũ Lão', '- Quận 1'), (b0['ma'], b0['tu']))
txt = [d['text'] for d in tl.dong(tl.pdf.pages[0])]
T('đọc đúng dòng có dấu khó', 'BIÊN BẢN VỀ SỰ CỐ, BẤT CẬP' in txt and 'HỆ THỐNG CHIẾU SÁNG ĐÔ THỊ' in txt and
  '2 - Người kiểm tra, phát hiện: Đào Thế Hiển' in txt)
T('không còn ký tự không đọc được trong mục 3', all('�' not in (b['tu'] + b['phuong']['text'] + b['quan']['text']) for b in bcs))

print('2. Tra tủ → phường mới')
tr = TraTu(rows)
T('trùng tên có dấu được ưu tiên (Nguyễn Văn Cừ 2 ≠ Nguyễn Văn Cự-2)', tr.tra('Nguyễn Văn Cừ 2', '1')[0] == 'Phường Cầu Ông Lãnh')
T('Nguyễn Trãi 6 (Q1) → Cầu Ông Lãnh, không lấy "Nguyễn Trãi - 6" Q5', tr.tra('Nguyễn Trãi 6', '1')[0] == 'Phường Cầu Ông Lãnh')
T('Điện Biên Phủ 9 → Tân Định', tr.tra('Điện Biên Phủ 9', '1')[0] == 'Phường Tân Định')
gia = TraTu([('Tủ X 1', 'Phường Bến Thành', '', '', ''), ('Tủ X 1', 'Phường Chợ Quán', '', '', '')])
T('tủ trùng tên 2 quận: chọn theo quận cũ', gia.tra('Tủ X 1', '1')[0] == 'Phường Bến Thành' and gia.tra('Tủ X 1', '5')[0] == 'Phường Chợ Quán')
p, g = gia.tra('Tủ X 1', '8')
T('trùng tên mà quận cũ không phân biệt được: không đổi, báo lại', p is None and 'nhiều phường' in g, g)
p, g = tr.tra('Nguyen Trai 6', '1')
T('chỉ khớp khi bỏ dấu (có thể đọc sai dấu): không tự đổi, báo lại', p is None and 'bỏ dấu' in g, g)
T('tủ không có trong bảng tra: không đổi, báo lại', tr.tra('Tủ Không Có 99', '1')[0] is None)

import pdfplumber, numpy as np
from PIL import Image
d = tempfile.mkdtemp()
moi_ph = None
def kiem_trang(fo, ten):
    """chu moi dung cho, cung co chu/chan chu; trang 1 chi khac trong vung 2 dong."""
    pm = pdfplumber.open(fo)
    T(ten + ': số trang giữ nguyên', len(pm.pages) == len(tl.pdf.pages))
    def chu_moi(i): return ''.join(c['text'] for c in pm.pages[i].chars if 'HCBCSC' in c['fontname'])
    T(ten + ': trang 1 có chữ mới đúng', chu_moi(0) == '- Phường Bến Thành- Thành phố Hồ Chí Minh', chu_moi(0))
    T(ten + ': mỗi trang báo cáo có đúng dòng phường mới', all(('- ' + moi_ph[b['trang'] + 1]) in chu_moi(b['trang']) for b in bcs))
    cm = [c for c in pm.pages[0].chars if 'HCBCSC' in c['fontname']]
    xm = [round(c['x0'], 1) for c in cm[:9]]; xg = [round(c['x0'], 1) for _, c in b0['phuong']['chars'][:9]]
    T(ten + ': độ rộng chữ đúng Times ("- Phường " trùng vị trí từng chữ với bản gốc)', all(abs(a - b) < 0.6 for a, b in zip(xm, xg)), (xm, xg))
    T(ten + ': cỡ chữ & đường chân chữ trùng bản gốc', all(abs(c['size'] - b0['phuong']['chars'][0][1]['size']) < 0.01 for c in cm) and
      abs(cm[0]['matrix'][5] - b0['phuong']['chars'][0][1]['matrix'][5]) < 0.01)
    if shutil.which('pdftoppm'):
        for f, n in ((PDF, 'a'), (fo, 'b')):
            subprocess.run(['pdftoppm', '-r', '72', '-f', '1', '-l', '1', '-png', f, os.path.join(d, n)], capture_output=True)
        A = np.asarray(Image.open(os.path.join(d, 'a-001.png')).convert('L'), int); B = np.asarray(Image.open(os.path.join(d, 'b-001.png')).convert('L'), int)
        m = np.abs(A - B) > 40; ys, xs = np.nonzero(m)
        y0 = min(b0['phuong']['top'], b0['quan']['top']) - 2; y1 = max(b0['phuong']['bottom'], b0['quan']['bottom']) + 3
        T(ten + ': trang 1 chỉ khác trong vùng 2 dòng phường/quận', m.any() and ys.min() >= y0 and ys.max() <= y1, (ys.min(), ys.max(), y0, y1))
    if shutil.which('qpdf'):
        q = subprocess.run(['qpdf', '--check', fo], capture_output=True, text=True)
        T(ten + ': qpdf --check không lỗi cấu trúc', 'No syntax or stream encoding errors' in q.stdout, q.stdout[-300:])
    return pm

print('3a. Xuất bản sạch (mặc định) — để ký lại')
r = chuyen(goc, rows, FONTS, log=lambda *a: None)
out = r['pdf']; moi_ph = {b['trang']: b['phuong_moi'] for b in r['bao_cao']}
T('đổi đủ 78/78', r['thong_ke'] == {'bao_cao': 78, 'da_doi': 78, 'khong_doi': 0}, r['thong_ke'])
T('bỏ 2 chữ ký cũ', r['chu_ky_bo'] == 2)
T('không còn lịch sử sửa (1 lần %%EOF, không chứa bản gốc)', out.count(b'%%EOF') == 1 and out[:len(goc)] != goc)
T('sạch: không sót dữ liệu chữ ký cũ (/ByteRange, chứng thư)', b'/ByteRange' not in out and b'adbe.pkcs7' not in out)
fo = os.path.join(d, 'sach.pdf'); open(fo, 'wb').write(out)
pm = kiem_trang(fo, 'sạch')
t1 = subprocess.run(['pdftotext', '-f', '1', '-l', '1', fo, '-'], capture_output=True, text=True) if shutil.which('pdftotext') else None
if t1:
    T('sạch: chữ cũ bị xoá hẳn (không còn "Phạm Ngũ Lão", "Quận 1")', 'Phạm Ngũ Lão' not in t1.stdout and 'Quận 1' not in t1.stdout and 'Phường Bến Thành' in t1.stdout)
    T('sạch: không còn tham chiếu hỏng', 'Syntax Error' not in t1.stderr, t1.stderr[:200])
    tt = subprocess.run(['pdftotext', fo, '-'], capture_output=True, text=True).stdout
    T('sạch: không còn dòng "- Quận" nào', '- Quận' not in tt and tt.count('- Thành phố Hồ Chí Minh') == 78)
if shutil.which('pdfsig'):
    s_ = subprocess.run(['pdfsig', fo], capture_output=True, text=True).stdout
    T('sạch: không còn ô ký / chữ ký số nào', 'Signature #' not in s_ and 'Signature is Valid' not in s_, s_[-300:])
from pypdf import PdfReader
T('sạch: đọc lại được bằng thư viện khác (pypdf)', len(PdfReader(fo).pages) == len(tl.pdf.pages))

print('3b. Ghi nối tiếp (giữ bản đã ký bên trong)')
r = chuyen(goc, rows, FONTS, log=lambda *a: None, kieu='noi_tiep')
out = r['pdf']
T('nối tiếp: file mới bắt đầu bằng nguyên byte file gốc', out[:len(goc)] == goc and len(out) > len(goc))
fo = os.path.join(d, 'noitiep.pdf'); open(fo, 'wb').write(out)
kiem_trang(fo, 'nối tiếp')
if shutil.which('pdfsig'):
    s_ = subprocess.run(['pdfsig', fo], capture_output=True, text=True).stdout
    T('nối tiếp: 2 chữ ký số bản gốc vẫn hợp lệ', s_.count('Signature is Valid') == 2, s_[-400:])

print('4. Ca xấu')
bot = [x for x in rows if x[0] != 'Nguyễn Du 2']
r2 = chuyen(goc, bot, FONTS, log=lambda *a: None)
loi = [b for b in r2['bao_cao'] if b['ghi_chu']]
T('thiếu 1 tủ trong bảng tra: báo 1 dòng, còn lại vẫn đổi', len(loi) == 1 and loi[0]['tu'] == 'Nguyễn Du 2' and r2['thong_ke']['da_doi'] == 77)
if KHAC:
    k = open(KHAC, 'rb').read(); r3 = chuyen(k, rows, FONTS, log=lambda *a: None)
    T('PDF không phải BCSC: 0 báo cáo, trả nguyên file', r3['thong_ke']['bao_cao'] == 0 and r3['pdf'] == k)
    r3 = chuyen(k, rows, FONTS, log=lambda *a: None, kieu='noi_tiep')
    T('PDF không phải BCSC (nối tiếp): trả nguyên file', r3['pdf'] == k)
try:
    chuyen(b'khong phai pdf', rows, FONTS, log=lambda *a: None); T('file hỏng: báo lỗi', False)
except Exception as e:
    T('file hỏng: báo lỗi rõ', True)
shutil.rmtree(d)
print('\nKết quả: %d/%d đạt' % (ok, ok + fail))
sys.exit(1 if fail else 0)
