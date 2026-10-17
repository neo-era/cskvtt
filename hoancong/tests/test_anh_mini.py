# -*- coding: utf-8 -*-
"""Gia lap trinh duyet (khong co Pillow): logo trang Bia phai con nguyen sau khi doc/ghi mau va co trong PDF.
python3 tests/test_anh_mini.py <mau .xlsm>"""
import sys, os, zipfile, tempfile, warnings
warnings.filterwarnings('ignore'); sys.modules['PIL'] = None
H = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, os.path.join(H, '..', 'src'))
import anh_mini; assert anh_mini.cai()
import openpyxl, pdfplumber, xuat_pdf as X
T = tempfile.mkdtemp(); src = sys.argv[1]
n0 = sum(i.file_size for i in zipfile.ZipFile(src).infolist() if '/media/' in i.filename)
wb = openpyxl.load_workbook(src); wb.save(T + '/o.xlsx')
n1 = sum(i.file_size for i in zipfile.ZipFile(T + '/o.xlsx').infolist() if '/media/' in i.filename)
R = X.xuat(T + '/o.xlsx', T + '/o.pdf', os.path.join(H, '..', 'fonts'), log=None)
ok1 = n1 == n0 > 0; ok2 = len(pdfplumber.open(T + '/o.pdf').pages[0].images) >= 1
print(('ĐẠT' if ok1 else 'LỖI') + '  logo giữ nguyên khi ghi Excel (%d / %d byte)' % (n1, n0))
print(('ĐẠT' if ok2 else 'LỖI') + '  logo có trong PDF trang Bìa')
sys.exit(0 if ok1 and ok2 else 1)
