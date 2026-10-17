# -*- coding: utf-8 -*-
"""Sinh src/do_rong_chu.py: be rong tung chu (Latin + tieng Viet + ky hieu) cua cac font thay the, de do/ngat dong
khong can tai file font (dung cho dinh dang in Excel). Chay lai khi doi font trong fonts/."""
import os, sys, json
H = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, os.path.join(H, '..', 'src'))
from xuat_pdf import TTF, FONT_FILE
rng = list(range(0x20, 0x7F)) + list(range(0xA0, 0x250)) + list(range(0x300, 0x370)) + list(range(0x1E00, 0x1F00)) + \
      list(range(0x2000, 0x2070)) + [0x2190, 0x2192, 0x2212, 0x2264, 0x2265, 0x2260, 0x221E, 0x2026, 0x20AB, 0x2116, 0x00B5, 0x03BC, 0x25CF, 0x2022]
out = {}
for k, f in FONT_FILE.items():
    p = os.path.join(H, '..', 'fonts', f)
    if not os.path.exists(p): continue
    t = TTF(open(p, 'rb').read(), f)
    chars = ''.join(chr(c) for c in rng if c in t.cmap)
    out['%s,%d,%d' % k] = {'upm': t.upm, 'asc': t.asc, 'desc': t.desc, 'ulpos': t.ulpos, 'ulth': t.ulth, 'cap': t.cap,
                           'w0': t.w[0], 'chars': chars, 'w': [t.w[t.cmap[ord(c)]] for c in chars]}
with open(os.path.join(H, '..', 'src', 'do_rong_chu.py'), 'w', encoding='utf-8') as fh:
    fh.write('# -*- coding: utf-8 -*-\n# SINH TU DONG boi tools/tao_do_rong_chu.py - KHONG sua tay. Be rong chu (don vi font) cua font thay the.\n')
    fh.write('M = ' + json.dumps(out, ensure_ascii=False, separators=(',', ':')) + '\n')
print({k: len(v['chars']) for k, v in out.items()}, os.path.getsize(os.path.join(H, '..', 'src', 'do_rong_chu.py')))
