# -*- coding: utf-8 -*-
"""Kiem chung tren BO BIEN BAN THAT (13 file, co bien ban 8-11): so hieu = bien ban, KH quy, nhom BD, ngay, vi tri tu.
python3 tests/test_bo_that.py <thu muc PDF/> <mau .xlsm> <bang tra .xlsx> <QUAN> <THANG>"""
import os, sys, re, glob, subprocess, tempfile, json, calendar, warnings
warnings.filterwarnings('ignore')
import openpyxl, pdfplumber
SRC = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'src')
D, MAU, MASTER, QUAN, THANG = sys.argv[1:6]; TMP = tempfile.mkdtemp(prefix='hc_that_')
OK, FAIL = [], []
def check(c, n, d=''): (OK if c else FAIL).append(n + ('' if c else '  <-- ' + str(d)))
out = os.path.join(TMP, 'o.xlsx')
env = dict(os.environ, SRC_DIR=D, MASTER_XLSX=MASTER, SRC_XLSM=MAU, GOI='G201NA', QUAN=QUAN, THANG=THANG, PREFER='', OUT_XLSX=out + '.raw.xlsx')
p = subprocess.run([sys.executable, '-c', "import runpy,report,json;G=runpy.run_path('refill_main.py');R=report.tieout(G);print('CK'+json.dumps(R['checks'],ensure_ascii=False))"], cwd=SRC, env=env, capture_output=True, text=True)
if p.returncode: print(p.stderr[-3000:]); raise SystemExit(1)
ck = json.loads([l for l in p.stdout.splitlines() if l.startswith('CK')][0][2:])
subprocess.run([sys.executable, 'dinhdang_in.py', out + '.raw.xlsx', out], cwd=SRC, check=True, capture_output=True)
for c in ck: check(c['ok'], 'bảng kiểm: ' + c['label'], c['detail'])
wb = openpyxl.load_workbook(out)
T = [(ws.title, c.coordinate, c.value) for ws in wb.worksheets for row in ws.iter_rows() for c in row if isinstance(c.value, str)]
# so hieu that tu bien ban 8-11
that = {}
for f in glob.glob(D + '*.pdf'):
    with pdfplumber.open(f) as pdf: t = ' '.join((pdf.pages[0].extract_text() or '').split())
    for m in re.finditer(r'(\d{2}/\d{2}/(?:YCNT|NTCV)-Q\d+)', t): that[m.group(1)] = 1
mine = set(m.group(1) for s, c, v in T if s[:2] in ('8.', '9.', '10', '11') for m in re.finditer(r'(\d{2}/\d{1,2}/(?:YCNT|NTCV)(?:-Q\d+)?)', v))
check(that and mine == set(that), 'số hiệu sheet 8–11 = biên bản 8–11', (sorted(mine), sorted(that)))
last = calendar.monthrange(2026, int(THANG))[1]
check(not [v for s, c, v in T if re.search(r'ngày\s+(\d+)\s+tháng\s+0?%s\s+năm' % THANG, v, re.I) and int(re.search(r'ngày\s+(\d+)', v, re.I).group(1)) > last], 'không có ngày vượt cuối tháng')
check(any(re.search(r'ngày\s+%d\s+tháng\s+0?%s\s+năm' % (last, THANG), v) for s, c, v in T), 'có ngày %d tháng %s' % (last, THANG))
noft = [w.title for w in wb.worksheets if w.sheet_state == 'visible' and not w.title.startswith('BIA') and 'Trang &P / &N' not in (w.oddFooter.right.text or '')]
check(not noft, 'số trang mọi sheet', noft)
ws6 = wb['6.HTCT_BD_G201']; grp = [ws6.cell(r, 2).value for r in range(33, ws6.max_row) if ws6.cell(r, 1).value is None and ws6.cell(r, 2).value and not str(ws6.cell(r, 2).value).startswith(('Ông', 'Bà'))]
check(len(grp) >= 2, 'sheet 6 có nhóm theo biên bản', grp[:5])
# sheet 8, 10 (phieu YCNT): sau bang phai con du muc 3. Thoi gian, 4. Dia diem, cau de nghi, roi moi den Chi huy truong
for sn in [w for w in wb.sheetnames if w[:2] in ('8.', '10')]:
    ws = wb[sn]; col = {}
    for r in range(33, ws.max_row + 1):
        for c in range(1, 6):
            v = ws.cell(r, c).value
            if isinstance(v, str):
                for k in ('Thời gian', 'Địa điểm', 'Kính đề nghị', 'Chỉ huy'):
                    if k in v and k not in col: col[k] = r
    check(len(col) == 4 and col['Thời gian'] < col['Địa điểm'] < col['Kính đề nghị'] < col['Chỉ huy'], sn + ': đủ mục 3. Thời gian, 4. Địa điểm, câu đề nghị trước chữ ký', col)
# nhan nhom (dong nhom) lay DUNG CHU bien ban, khong dung chu cua mau
nhom_pdf = set()
for f in glob.glob(D + '*.pdf'):
    with pdfplumber.open(f) as pdf:
        for pg in pdf.pages[:2]:
            for t in pg.extract_tables():
                for r in t:
                    r = [' '.join((c or '').split()) for c in r]
                    if len(r) > 3 and (r[0] == '' or re.fullmatch(r'[IVX]+', r[0])) and not r[3] and re.match(r'(Tủ |Công tác duy tu)', r[1]): nhom_pdf.add(' '.join((r[1] + r[2] if not r[3] and len(r[2]) < 6 else r[1]).split()))
nhom_pdf = {x.replace('vê ', 'về ') for x in nhom_pdf}
ov = set(v.strip() for s, c, v in T if s[:2] in ('2.', '3.', '4_', '5.', '6.', '7_', '8.', '9.', '10', '11'))
mau_only = [x for x in ('Tủ có kết nối', 'Tủ thường', 'Bảo dưỡng thường xuyên', 'Bảo dưỡng không thường xuyên') if x in ov and x not in nhom_pdf]
check(not mau_only, 'nhãn nhóm không còn chữ của mẫu', mau_only)
for sn, k in (('4_', 'Tủ kết nối về trung tâm'), ('3.', 'Tủ kết nối về trung tâm'), ('2.', 'Tủ kết nối về trung tâm'), ('8.', 'Tủ kết nối về trung tâm'), ('6.', 'Công tác duy tu thường xuyên'), ('10', 'Công tác duy tu không thường xuyên')):
    if k in nhom_pdf: check(any(s.startswith(sn) and v.strip() == k for s, c, v in T), 'sheet %s có nhãn nhóm "%s" như biên bản' % (sn, k))
mat = [(s, c, v[:50]) for s, c, v in T if re.search(r'(?<!\w)(vê|điêu|nhiêu|đêu)(?!\w)', v, re.I)]
check(not mat, 'không còn chữ mất dấu huyền (vê, điêu…) do đọc PDF', mat[:3])
# nhan Goi thau / Du toan mua sam / Dia diem: dam + gach duoi, phan noi dung chu thuong; chu giu nguyen
wr = openpyxl.load_workbook(out, rich_text=True); nl = []; nn = 0
for ws in wr.worksheets:
    for row in ws.iter_rows(max_row=40):
        for c in row:
            v = c.value; t = str(v) if v is not None else ''
            if re.match(r'\s*(Gói thầu|Dự toán mua sắm|Địa điểm)\s*:\s*\S', t):
                nn += 1; b = list(v) if type(v).__name__ == 'CellRichText' else []
                ok = len(b) >= 2 and getattr(b[0], 'font', None) is not None and b[0].font.u == 'single' and b[0].font.b and not b[-1].font.b
                if not ok: nl.append((ws.title, c.coordinate))
check(nn >= 20 and not nl, 'nhãn Gói thầu/Dự toán mua sắm/Địa điểm: đậm + gạch dưới, nội dung chữ thường (%d ô)' % nn, nl[:5])
mt = openpyxl.load_workbook(MAU); diff = []
for ws in wr.worksheets:
    if ws.title not in mt.sheetnames: continue
    for row in ws.iter_rows(max_row=14):
        for c in row:
            t = str(c.value) if c.value is not None else ''
            if re.match(r'\s*(Gói thầu|Dự toán mua sắm)\s*:', t) and t != str(mt[ws.title][c.coordinate].value): diff.append((ws.title, c.coordinate))
check(not diff, 'chữ Gói thầu / Dự toán mua sắm giữ nguyên như mẫu', diff[:3])
pa4 = [w.print_area for w in wb.worksheets if w.title.startswith('4_')]
check(pa4 and re.search(r'\$C\$1:\$K\$\d+', pa4[0]), 'sheet 4: vùng in C:K như mẫu (mẫu ghi dạng chỉ cột $C:$K)', pa4)
print('\n'.join('ĐẠT  ' + x for x in OK)); print('\n'.join('LỖI  ' + x for x in FAIL)); print('Tổng: %d đạt, %d lỗi' % (len(OK), len(FAIL)))
print('FILE', out)
sys.exit(1 if FAIL else 0)
