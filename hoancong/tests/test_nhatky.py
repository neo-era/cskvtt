# -*- coding: utf-8 -*-
"""Kiem thu tab Nhat ky thi cong (src/nhatky.py).
Chay: python3 tests/test_nhatky.py <NKTC.pdf (Crystal, nhieu ngay)> <Ma cong viec .xlsb> [BCSC.pdf de thu ca xau]
Du lieu that KHONG dua len repo (repo cong khai)."""
import io, json, math, os, re, subprocess, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', 'src'))
FONTS = os.path.join(HERE, '..', 'fonts')
NK, XLSB = sys.argv[1], sys.argv[2]
BCSC = sys.argv[3] if len(sys.argv) > 3 else None
ok = fail = 0


def T(ten, dk, chi_tiet=''):
    global ok, fail
    if dk: ok += 1; print('  ĐẠT ', ten)
    else: fail += 1; print('  LỖI ', ten, chi_tiet)


import nhatky as N

print('1. Đọc định mức (Mã công việc)')
dm = N.doc_dinh_muc(open(XLSB, 'rb').read(), os.path.basename(XLSB))
T('đọc được nhiều mục định mức', len(dm['muc']) >= 500, len(dm['muc']))
m = N.tra_dinh_muc(dm, 'Tháo dỡ cáp nổi')
T('Tháo dỡ cáp nổi: 0,626 công + xe thang 9m 0,124 ca / 100m', m and abs(m['nc'] - 0.626) < 1e-9 and abs(m['may'].get('Xe thang 9m', 0) - 0.124) < 1e-9 and m['dv'] == '100m', m)
T('tra không phân biệt dấu/hoa thường/khoảng trắng', N.tra_dinh_muc(dm, '  tháo  dỡ CÁP nổi ') == m)
T('định mức gọn để lưu trên máy (JSON < 400 KB)', len(json.dumps(dm, ensure_ascii=False)) < 400000)

print('2. Số khối lượng kiểu Việt Nam')
T('0,560 → 0.56', N.so_kl('0,560') == 0.56)
T('20,000 → 20', N.so_kl('20,000') == 20)
T('1.200,500 → 1200.5', N.so_kl('1.200,500') == 1200.5)

print('3. Nhân lực, thiết bị theo định mức (tối thiểu 1 tổ: 04 người, 01 xe thang)')
pn = [('Cung cấp kẹp dừng cáp', 6, 'cái'), ('Lắp đặt cáp nguồn nổi', 0.3, '100m'), ('Lắp kẹp đầu cáp', 20, 'cái'), ('Tháo dỡ cáp nổi', 0.56, '100m')]
k = N.tinh_nhan_luc(pn, dm)
T('ngày 23/10 PN (ảnh mẫu): 04 người, 01 xe thang', k['cong_nhan'] == 4 and k['thiet_bi'] == '01 xe thang', k)
T('"Lắp kẹp đầu cáp" khớp định mức "Lắp kẹp đấu cáp"', not k['thieu'], k['thieu'])
T('tổng công tính đúng (≈2,91)', abs(k['cong'] - (0.3 * 1.63 + 20 * 0.103 + 0.56 * 0.626)) < 1e-6, k['cong'])
nhieu = [('Lắp kẹp đầu cáp', 70, 'cái')]
k2 = N.tinh_nhan_luc(nhieu, dm)
T('khối lượng lớn: số người = làm tròn lên tổng công (70×0,103 → 08)', k2['cong_nhan'] == math.ceil(70 * 0.103), k2)
T('xe thang làm tròn lên theo tổng ca (70×0,026 → 02)', k2['thiet_bi'].startswith('02 xe thang'), k2['thiet_bi'])
k5 = N.tinh_nhan_luc([('Lắp đặt cáp nguồn nổi', 9.91, '100m'), ('Lắp kẹp đầu cáp', 122, 'cái')], dm)
T('vượt năng lực (Q8 19/10): chặn ở 12 người, 03 xe thang và báo vượt', k5['cong_nhan'] == 12 and k5['thiet_bi'].startswith('03 xe thang') and k5['vuot'] and k5['cong_nhan_dm'] > 12 and k5['xe_dm'] > 3, k5)
k6 = N.tinh_nhan_luc([('Cung cấp gói data 3G/4G kết nối tủ về trung tâm điều khiển', 1, 'tủ')], dm)
T('ngày chỉ cung cấp vật tư/dịch vụ: để trống nhân lực, thiết bị', k6['chi_cung_cap'] and k6['cong_nhan'] == 0 and k6['thiet_bi'] == '', k6)
k7 = N.tinh_nhan_luc([('Lắp đặt cáp nguồn nổi', 0.5, '100m'), ('Việc lạ', 1, 'cái')], dm, tran=(0, 0))
T('không chặn khi trần = 0', not k7['vuot'] and k7['cong_nhan'] == 4, k7)
k3 = N.tinh_nhan_luc([('Việc lạ chưa có định mức', 3, 'cái')], dm)
T('việc chưa có định mức: báo thiếu, vẫn tối thiểu 1 tổ', k3['thieu'] == ['Việc lạ chưa có định mức'] and k3['cong_nhan'] == 4 and k3['thiet_bi'] == '01 xe thang', k3)
k4 = N.tinh_nhan_luc([], dm)
T('ngày không có việc: vẫn 1 tổ', k4['cong_nhan'] == 4 and k4['thiet_bi'] == '01 xe thang', k4)

print('4. Thời tiết theo giờ (Open-Meteo) → sáng 7–11h, chiều 13–17h')
gio = ['2026-10-23T%02d:00' % h for h in range(24)]
nd = [26 + (h - 6) * 0.5 if 6 <= h <= 14 else 30 for h in range(24)]
mua = [0] * 24; mua[15] = 2.4
tt = N.thoi_tiet({'hourly': {'time': gio, 'temperature_2m': nd, 'precipitation': mua}})
d = tt['2026-10-23']
T('sáng: min–max 7..11h, làm tròn', d['sang'] == '27°C – 29°C', d)
T('chiều: min–max 13..17h', d['chieu'] == '30°C – 30°C' or d['chieu'] == '30°C', d)
T('mưa buổi chiều → "Có mưa", sáng vẫn bình thường', d['mua_chieu'] and not d['mua_sang'], d)
T('dữ liệu rỗng/lỗi → không có ngày nào, không sập', N.thoi_tiet({}) == {} and N.thoi_tiet({'hourly': {'time': gio, 'temperature_2m': [None] * 24, 'precipitation': [None] * 24}}) == {})

print('5. Đọc file nhật ký (NKTC)')
goc = open(NK, 'rb').read()
ngay = N.doc_nhat_ky(goc)
import pdfplumber
pp = pdfplumber.open(io.BytesIO(goc))
so_ngay = sum(1 for p in pp.pages if 'NHẬT KÝ THI CÔNG HÀNG NGÀY' in (p.extract_text() or ''))
T('đủ số ngày như số tiêu đề trong file', len(ngay) == so_ngay and so_ngay > 5, (len(ngay), so_ngay))
T('mỗi trang thuộc đúng một ngày', sorted(t for n_ in ngay for t in n_['trang']) == list(range(len(pp.pages))))
d1 = ngay[0]
T('ngày đầu: 2026-10-01, trang 1–3', d1['ngay'] == '2026-10-01' and d1['trang'] == [0, 1, 2], (d1['ngay'], d1['trang']))
mu = {x['ten']: x for x in d1['muc']}
T('việc lặp ở đầu trang sau không cộng hai lần', abs(mu['Thay thế bóng, chấn lưu, kích đèn HPS 2 cấp công suất 150/100W']['kl'] - 7) < 1e-9, mu.get('Thay thế bóng, chấn lưu, kích đèn HPS 2 cấp công suất 150/100W'))
T('đủ vị trí điền cho mọi ngày (nhiệt độ, công nhân, thiết bị, 2×3 ô tick)',
  all(n_['vi_tri'].get(k_) for n_ in ngay for k_ in ('nhiet_do', 'cong_nhan', 'thiet_bi', 'vsmt', 'atld')) and all(len(n_['vi_tri']['vsmt']) == 3 == len(n_['vi_tri']['atld']) for n_ in ngay))
ng_ngat = [n_ for n_ in ngay if n_['vi_tri']['vsmt'][0]['trang'] != [p_ for p_ in n_['trang'] if any('4. Công tác vệ sinh' in (l_['text']) for l_ in pp.pages[p_].extract_text_lines())][0]]
T('mục 4 ngắt trang (ô tick sang trang sau) vẫn tìm đúng ô', bool(ng_ngat) and all(n_['vi_tri']['vsmt'][0]['trang'] < n_['vi_tri']['atld'][0]['trang'] or n_['vi_tri']['vsmt'][0]['top'] < n_['vi_tri']['atld'][0]['top'] for n_ in ngay), [n_['ngay'] for n_ in ng_ngat])
de_xuat = N.de_xuat(ngay, dm)
T('đề xuất đủ mọi ngày, không thiếu định mức với file thật', len(de_xuat) == len(ngay) and not any(x['thieu'] for x in de_xuat), [x['thieu'] for x in de_xuat if x['thieu']][:3])
T('đề xuất mặc định tick "Bình thường"', all(x['vsmt'] == 'binh_thuong' == x['atld'] for x in de_xuat))

loi = [(n_['ngay'], t_) for n_ in ngay for t_ in n_['doc_loi']]
T('mọi dòng "•" đều đọc được (kể cả đơn vị nhiều chữ "đầu cáp", "1 Trụ")', not loi, loi[:4])
d7 = [n_ for n_ in ngay if n_['ngay'] == '2026-10-07'][0]
dv7 = {x['ten']: x['dv'] for x in d7['muc']}
T('ngày 07/10: đủ 21 việc, có "Tháo dỡ trụ đèn thép" đơn vị "1 Trụ"', len(d7['muc']) == 21 and dv7.get('Tháo dỡ trụ đèn thép') == '1 Trụ', (len(d7['muc']), dv7.get('Tháo dỡ trụ đèn thép')))
T('không ngày nào thiếu ô cần điền', not any(n_['thieu_o'] for n_ in ngay), [(n_['ngay'], n_['thieu_o']) for n_ in ngay if n_['thieu_o']][:3])
T('ngày 30/10 chỉ cung cấp gói data → để trống nhân lực', [x for x in de_xuat if x['ngay'] == '2026-10-30'][0]['chi_cung_cap'])

print('5b. Ca giả: dòng • hỏng, tên việc xuống dòng, thiếu chỗ điền')
from reportlab.pdfgen import canvas as _cv
from reportlab.pdfbase import pdfmetrics as _pm
from reportlab.pdfbase.ttfonts import TTFont as _TT
_pm.registerFont(_TT('L', os.path.join(FONTS, 'LiberationSerif-Regular.ttf')))
_b = io.BytesIO(); _c = _cv.Canvas(_b); _c.setFont('L', 11); _y = 800
for _s in ['NHẬT KÝ THI CÔNG HÀNG NGÀY', 'Ngày 05 tháng 10 năm 2026', '1. Công việc',
           '• Lắp đặt cáp nguồn nổi: 0,500 100m', '• Thay thế bóng, chấn lưu, kích đèn HPS 2 cấp công suất',
           '150/100W: 3,000 bộ', '• Việc hỏng không có khối lượng', '- Ghi chú', '• Tên: abc cái']:
    _c.drawString(40, _y, _s); _y -= 18
_c.save()
_n = N.doc_nhat_ky(_b.getvalue())[0]
_mu = {m_['ten']: (m_['kl'], m_['dv']) for m_ in _n['muc']}
T('tên việc xuống dòng được nối lại', _mu.get('Thay thế bóng, chấn lưu, kích đèn HPS 2 cấp công suất 150/100W') == (3.0, 'bộ'), _mu)
T('dòng • không đọc được thì báo trong doc_loi (2 dòng)', len(_n['doc_loi']) == 2 and any('Việc hỏng' in x for x in _n['doc_loi']), _n['doc_loi'])
T('ngày thiếu chỗ điền được báo (thieu_o đủ 6 mục)', len(_n['thieu_o']) == 6, _n['thieu_o'])
_dx = N.de_xuat([_n], dm)[0]
T('đề xuất mang theo cảnh báo doc_loi/thieu_o để app tô đỏ', _dx['doc_loi'] and _dx['thieu_o'])

print('6. Điền vào PDF')
gt = {}
for i, x in enumerate(de_xuat):
    g = dict(x); g['sang'] = '29°C – 31°C'; g['chieu'] = '28°C – 32°C'
    if i == 1: g['mua_chieu'] = True
    if i == 2: g['vsmt'] = 'tot'; g['atld'] = 'kem'
    gt[x['ngay']] = g
R = N.dien(goc, gt, FONTS, log=lambda *a: None)
out = R['pdf']
po = pdfplumber.open(io.BytesIO(out))
T('giữ nguyên số trang', len(po.pages) == len(pp.pages))
t1 = po.pages[0].extract_text()
T('nhiệt độ sáng/chiều đúng dòng "Nhiệt độ"', re.search(r'Nhiệt độ\s+29°C – 31°C\s+28°C – 32°C', t1) is not None, t1[:400])
T('số công nhân điền trước chữ "người"', re.search(r'Công nhân:\s*%s\s*người' % r'\s*'.join('%02d' % de_xuat[0]['cong_nhan']), t1) is not None, re.findall(r'Công nhân:.*', t1))
T('thiết bị điền sau "Thiết bị thi công:"', re.search(r'Thiết bị thi công:\s*' + re.escape(de_xuat[0]['thiet_bi']), t1) is not None)
p2 = po.pages[ngay[1]['trang'][0]].extract_text()
T('ngày có mưa chiều: "Bình thường  Có mưa", không còn chữ cũ chồng lên', re.search(r'Thời tiết\s+Bình thường\s+Có mưa\s*$', p2, re.M) is not None, [l for l in p2.splitlines() if 'Thời tiết' in l])
T('trang không có ô cần điền giữ nguyên nội dung chữ', po.pages[1].extract_text() == pp.pages[1].extract_text() or ngay[0]['vi_tri']['vsmt'][0]['trang'] == 1)
T('báo số ô đã tick = 2 × số ngày', R['tick'] == 2 * len(ngay), (R['tick'], len(ngay)))
with tempfile.TemporaryDirectory() as d:
    f = os.path.join(d, 'o.pdf'); open(f, 'wb').write(out)
    q = subprocess.run(['qpdf', '--check', f], capture_output=True, text=True)
    T('qpdf --check không lỗi', q.returncode == 0 and 'No syntax or stream encoding errors found' in q.stdout, q.stdout[-300:])
    # o tick: so pixel muc dam trong o "Tot" ngay 3 tang so voi ban goc
    vt = ngay[2]['vi_tri']['vsmt'][0]; tr = vt['trang'] + 1
    g0 = os.path.join(d, 'g'); g1 = os.path.join(d, 'h')
    fg = os.path.join(d, 'goc.pdf'); open(fg, 'wb').write(goc)
    subprocess.run(['pdftoppm', '-r', '72', '-gray', '-f', str(tr), '-l', str(tr), fg, g0], check=True)
    subprocess.run(['pdftoppm', '-r', '72', '-gray', '-f', str(tr), '-l', str(tr), f, g1], check=True)
    from PIL import Image
    a = Image.open([os.path.join(d, x) for x in sorted(os.listdir(d)) if x.startswith('g-')][0])
    b = Image.open([os.path.join(d, x) for x in sorted(os.listdir(d)) if x.startswith('h-')][0])
    box = (int(vt['x0']), int(vt['top']), int(vt['x1']) + 1, int(vt['bottom']) + 1)
    den = lambda im: sum(1 for v in im.crop(box).getdata() if v < 128)
    T('ô "Tốt" (ngày 3) có dấu tick vẽ vào', den(b) > den(a) + 6, (den(a), den(b)))
    vk = ngay[2]['vi_tri']['vsmt'][1]; box = (int(vk['x0']), int(vk['top']), int(vk['x1']) + 1, int(vk['bottom']) + 1)
    T('ô "Bình thường" ngày 3 không bị tick', den(b) <= den(a) + 1, (den(a), den(b)))
rd = __import__('pypdf').PdfReader(io.BytesIO(out))
T('bỏ chữ ký (mặc định): không còn ô ký số', not any((a.get_object().get('/FT') == '/Sig') for p in rd.pages for a in (p.get('/Annots') or [])) and b'/ByteRange' not in out)

print('7. Ca xấu')
try:
    N.doc_nhat_ky(b'%PDF-1.4 hong'); T('PDF hỏng → báo lỗi rõ', False)
except ValueError as e: T('PDF hỏng → báo lỗi rõ', 'PDF' in str(e), e)
if BCSC:
    try:
        N.doc_nhat_ky(open(BCSC, 'rb').read()); T('file không phải nhật ký → báo lỗi', False)
    except ValueError as e: T('file không phải nhật ký → báo lỗi', 'nhật ký' in str(e).lower(), e)
try:
    N.doc_dinh_muc(b'PK\x03\x04 khong phai', 'x.xlsb'); T('file định mức hỏng → báo lỗi rõ', False)
except ValueError as e: T('file định mức hỏng → báo lỗi rõ', 'định mức' in str(e).lower(), e)
gt2 = {k_: dict(v, cong_nhan=0, thiet_bi='') for k_, v in gt.items()}
R2 = N.dien(goc, gt2, FONTS, log=lambda *a: None)
T('anh xoá trắng công nhân/thiết bị → để trống, không ghi "00"', '00 ' not in pdfplumber.open(io.BytesIO(R2['pdf'])).pages[0].extract_text().split('Công nhân:')[1][:12])

print('\nKết quả: %d/%d đạt' % (ok, ok + fail))
sys.exit(1 if fail else 0)
