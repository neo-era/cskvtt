# -*- coding: utf-8 -*-
"""Xuat ho so .xlsx -> 1 file PDF, giong ban in Excel (A4, le, vung in, vua 1 trang ngang, lap tieu de, chan trang).
Thuan Python (chay ca tren trinh duyet qua Pyodide): tu tinh cong thuc SUM..., dinh dang so kieu Viet Nam,
dan trang, ve PDF va nhung font (cat gon chi giu chu dung den).
   python3 xuat_pdf.py HO_SO.xlsx HO_SO.pdf [thu_muc_font]
Font thay the cung kho chu: Times New Roman -> Liberation Serif; Calibri -> Carlito (xem fonts/LICENSE)."""
import os, re, io, sys, math, zlib, struct, time, datetime, warnings, unicodedata
from decimal import Decimal, ROUND_HALF_UP
warnings.filterwarnings('ignore')
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import column_index_from_string as CI, get_column_letter as CL

# =====================================================================================
# 1. DINH DANG SO (kieu Viet Nam: phan nghin '.', phan thap phan ',')
# =====================================================================================
TP, NGHIN = ',', '.'

def _tach_muc(fmt):
    out, cur, q, i = [], '', False, 0
    while i < len(fmt):
        ch = fmt[i]
        if ch == '"': q = not q
        if ch == '\\' and i + 1 < len(fmt): cur += fmt[i:i + 2]; i += 2; continue
        if ch == ';' and not q: out.append(cur); cur = ''
        else: cur += ch
        i += 1
    out.append(cur)
    return out

def _so_general(v):
    if isinstance(v, int) or (isinstance(v, float) and v.is_integer() and abs(v) < 1e15):
        return str(int(v))
    a = abs(v)
    if a >= 1e11 or a < 1e-9:
        s = '%.5E' % v; m, e = s.split('E'); m = m.rstrip('0').rstrip('.')
        return m.replace('.', TP) + 'E' + ('+' if int(e) >= 0 else '-') + '%02d' % abs(int(e))
    nd = max(0, 10 - len(str(int(a))))
    s = ('%.' + str(nd) + 'f') % v
    if '.' in s: s = s.rstrip('0').rstrip('.')
    return s.replace('.', TP)

def _muc_so(v, sec):
    """Ap 1 muc dinh dang cho so v (v da la tri tuyet doi neu muc am co dau rieng)."""
    sec = re.sub(r'\[[^\]]*\]', '', sec)                    # [Red], [$-42A]...
    out = []; i = 0; ph = None                              # ph: vi tri chen phan so
    toks = []
    while i < len(sec):
        ch = sec[i]
        if ch == '"':
            j = sec.find('"', i + 1); j = len(sec) if j < 0 else j
            toks.append(('t', sec[i + 1:j])); i = j + 1; continue
        if ch == '\\' and i + 1 < len(sec): toks.append(('t', sec[i + 1])); i += 2; continue
        if ch == '_' and i + 1 < len(sec): toks.append(('t', ' ')); i += 2; continue
        if ch == '*' and i + 1 < len(sec): i += 2; continue
        if ch in '0#?.,%':
            j = i
            while j < len(sec) and sec[j] in '0#?.,%': j += 1
            toks.append(('n', sec[i:j])); i = j; continue
        if ch == '@': toks.append(('@', '')); i += 1; continue
        toks.append(('t', ch)); i += 1
    num = ''.join(t for k, t in toks if k == 'n')
    if '%' in num: v = v * 100
    if num.replace('%', '') == '':
        return ''.join(t for k, t in toks if k == 't').strip('') or ''
    ip, _, dp = num.replace('%', '').partition('.')
    nghin = ',' in ip
    ip2 = ip.replace(',', '')
    nd = sum(1 for c in dp if c in '0#?')
    if not any(c in '0#' for c in ip2 + dp):              # chi co '?' (vd "-"??) -> khoang trang
        s = ''
    else:
        d = Decimal(repr(float(v))).quantize(Decimal(1).scaleb(-nd), rounding=ROUND_HALF_UP) if nd else Decimal(repr(float(v))).quantize(Decimal(1), rounding=ROUND_HALF_UP)
        s = format(abs(d), 'f'); neg = d < 0
        a, _, b = s.partition('.')
        mi = ip2.count('0')
        if a == '0' and mi == 0: a = ''
        a = a.rjust(mi, '0')
        if nghin and a:
            a = '{:,}'.format(int(a)).replace(',', NGHIN) if a.isdigit() else a
        # chu so thap phan: '#' bo so 0 cuoi
        if nd:
            keep = sum(1 for c in dp if c == '0'); b = b[:nd]
            while len(b) > keep and b.endswith('0'): b = b[:-1]
        s = ('-' if neg else '') + a + ((TP + b) if nd and b else ('' if not dp else ''))
    res = []; put = False
    for k, t in toks:
        if k == 'n':
            if not put: res.append(s); put = True
        elif k == 't': res.append(t)
    if '%' in num: res.append('%')
    return ''.join(res)

def dinh_dang(v, fmt='General'):
    """Gia tri -> chu hien thi nhu Excel (kieu Viet Nam)."""
    if v is None: return ''
    fmt = fmt or 'General'
    if isinstance(v, bool): return 'TRUE' if v else 'FALSE'
    if isinstance(v, (datetime.datetime, datetime.date, datetime.time)):
        return _ngay(v, fmt)
    if isinstance(v, str):
        secs = _tach_muc(fmt)
        return v
    if fmt in ('General', '@') or fmt.lower() == 'general': return _so_general(v)
    secs = _tach_muc(fmt)
    if v > 0 or len(secs) == 1: sec = secs[0]; val = v
    elif v < 0: sec = secs[1] if len(secs) > 1 else secs[0]; val = abs(v) if len(secs) > 1 else v
    else: sec = secs[2] if len(secs) > 2 else secs[0]; val = 0
    if re.search(r'(?i)general', sec):                                 # vd "Tổng: "General
        return re.sub(r'(?i)general', lambda m: _so_general(val), re.sub(r'"([^"]*)"', r'\1', sec), count=1)
    me = re.search(r'([0#.,]+)E([+-])([0#]+)', sec)
    if me and val:                                                     # khoa hoc: 0.00E+00
        nd = me.group(1).partition('.')[2].count('0'); e = int(math.floor(math.log10(abs(val)))); m = val / 10 ** e
        if round(abs(m), nd) >= 10: m /= 10; e += 1
        txt = ('%.' + str(nd) + 'f') % m
        return sec[:me.start()] + txt.replace('.', TP) + 'E' + ('-' if e < 0 else '+') + str(abs(e)).zfill(len(me.group(3))) + sec[me.end():]
    mk = re.search(r'[0#](,+)(?=[^0#,]*$)', sec.split('.')[0] if '.' in sec else sec)
    if mk and '"' not in sec[:mk.end()]:                                 # #,##0, -> chia 1000
        val = val / (1000 ** len(mk.group(1))); sec = sec[:mk.start() + 1] + sec[mk.end():]
    return _muc_so(val, sec)

_THU = ['Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy', 'Chủ Nhật']
def _ngay(v, fmt):
    f = re.sub(r'\[[^\]]*\]', '', fmt or '')
    if isinstance(v, datetime.time): v = datetime.datetime(1900, 1, 1, v.hour, v.minute, v.second)
    if not isinstance(v, datetime.datetime): v = datetime.datetime(v.year, v.month, v.day)
    if not re.search(r'[dmyhs]', f, re.I): f = 'dd/mm/yyyy' if (v.hour, v.minute, v.second) == (0, 0, 0) else 'dd/mm/yyyy hh:mm'
    toks = re.findall(r'"[^"]*"|\\.|yyyy|yy|d{1,4}|m{1,5}|h{1,2}|s{1,2}|AM/PM|am/pm|.', f, re.I)
    out = []; pm = any(t.upper() == 'AM/PM' for t in toks)
    for i, t in enumerate(toks):
        tl = t.lower()
        if t.startswith('"'): out.append(t[1:-1]); continue
        if t.startswith('\\'): out.append(t[1:]); continue
        if tl == 'yyyy': out.append('%04d' % v.year)
        elif tl == 'yy': out.append('%02d' % (v.year % 100))
        elif tl in ('d', 'dd'): out.append(('%02d' if tl == 'dd' else '%d') % v.day)
        elif tl in ('ddd', 'dddd'): out.append(_THU[v.weekday()])
        elif tl[0] == 'm':
            gio = any(x.lower().startswith('h') for x in toks[max(0, i - 3):i]) or any(x.lower().startswith('s') for x in toks[i + 1:i + 4])
            if gio and len(tl) <= 2: out.append(('%02d' if tl == 'mm' else '%d') % v.minute)
            elif len(tl) <= 2: out.append(('%02d' if tl == 'mm' else '%d') % v.month)
            else: out.append('Tháng %d' % v.month if len(tl) == 4 else 'Th%d' % v.month)
        elif tl[0] == 'h':
            hh = v.hour % 12 or 12 if pm else v.hour; out.append(('%02d' if tl == 'hh' else '%d') % hh)
        elif tl[0] == 's': out.append(('%02d' if tl == 'ss' else '%d') % v.second)
        elif tl == 'am/pm': out.append('SA' if v.hour < 12 else 'CH')
        else: out.append(t)
    return ''.join(out)

# =====================================================================================
# 2. TINH CONG THUC (SUM, ROUND, + - * /, tham chieu o / vung / sheet khac)
# =====================================================================================
class KhongTinhDuoc(Exception): pass
class _QuaSau(Exception): pass          # chuoi tham chieu qua sau: hoan, tinh o luot sau
_TOK = re.compile(r"""\s*(?:
  (?P<str>"(?:[^"]|"")*")
 |(?P<bool>TRUE|FALSE)(?![A-Z0-9(])
 |(?P<ref>(?:(?:'(?:[^']|'')+'|[A-Za-z0-9_.À-ỹ]+)!)?(?:\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}(?![A-Z0-9])))
 |(?P<fn>[A-Z][A-Z0-9.]*)\(
 |(?P<num>\d+(?:\.\d*)?(?:[eE][+-]?\d+)?|\.\d+)
 |(?P<op><>|<=|>=|[-+*/^&=<>(),:%])
)""", re.X)

class TinhCongThuc:
    def __init__(self, wb):
        self.wb = wb; self.cache = {}; self.dang = set(); self.loi = []; self.luot = False
    def gia_tri(self, sh, r, c):
        k = (sh, r, c)
        if k in self.cache: return self.cache[k]
        v = self.wb[sh].cell(r, c).value
        if isinstance(v, str) and v.startswith('='):
            if k in self.dang: return 0
            if len(self.dang) > 120 and self.luot: raise _QuaSau()
            self.dang.add(k)
            try: v = self._tinh(v[1:], sh)
            except _QuaSau:
                self.dang.discard(k); raise
            except RecursionError:
                self.dang.discard(k)
                if self.luot: raise _QuaSau()
                raise
            except KhongTinhDuoc as e:
                self.loi.append('%s!%s%d: =%s (%s)' % (sh, CL(c), r, str(self.wb[sh].cell(r, c).value)[1:60], e)); v = None
            except Exception as e:
                self.loi.append('%s!%s%d: =%s (lỗi %s)' % (sh, CL(c), r, str(self.wb[sh].cell(r, c).value)[1:60], e)); v = None
            self.dang.discard(k)
        elif v is not None and type(v).__name__ == 'CellRichText': v = str(v)
        self.cache[k] = v
        return v
    # ---- phan tich cu phap
    def _tinh(self, s, sh):
        toks = []; pos = 0
        while pos < len(s):
            m = _TOK.match(s, pos)
            if not m or m.end() == pos:
                if s[pos:].strip() == '': break
                raise KhongTinhDuoc('không đọc được "%s"' % s[pos:pos + 15])
            toks.append((m.lastgroup, m.group(m.lastgroup))); pos = m.end()
        luu = (getattr(self, 't', None), getattr(self, 'i', 0), getattr(self, 'sh', None), getattr(self, 'bo', False))   # cong thuc long nhau
        self.t = toks; self.i = 0; self.sh = sh; self.bo = False
        try:
            v = self._cmp()
            if self.i != len(self.t): raise KhongTinhDuoc('thừa "%s"' % self.t[self.i][1])
            return v
        finally:
            self.t, self.i, self.sh, self.bo = luu
    def _peek(self): return self.t[self.i] if self.i < len(self.t) else (None, None)
    def _eat(self): x = self.t[self.i]; self.i += 1; return x
    def _cmp(self):
        a = self._add()
        while self._peek()[1] in ('=', '<>', '<', '>', '<=', '>='):
            op = self._eat()[1]; b = self._add()
            if self.bo: a = 0; continue
            a, b = self._sc(a), self._sc(b)
            a = {'=': a == b, '<>': a != b, '<': a < b, '>': a > b, '<=': a <= b, '>=': a >= b}[op]
        return a
    def _add(self):
        a = self._mul()
        while self._peek()[1] in ('+', '-', '&'):
            op = self._eat()[1]; b = self._mul()
            if self.bo: a = 0; continue
            if op == '&': a = dinh_dang(self._sc(a)) + dinh_dang(self._sc(b)) if False else _txt(self._sc(a)) + _txt(self._sc(b))
            else: a = _n(self._sc(a)) + (_n(self._sc(b)) if op == '+' else -_n(self._sc(b)))
        return a
    def _mul(self):
        a = self._pow()
        while self._peek()[1] in ('*', '/'):
            op = self._eat()[1]; b = self._pow()
            if self.bo: a = 0; continue
            b = _n(self._sc(b))
            if op == '/' and b == 0: raise KhongTinhDuoc('chia cho 0')
            a = _n(self._sc(a)) * b if op == '*' else _n(self._sc(a)) / b
        return a
    def _pow(self):
        a = self._un()
        while self._peek()[1] == '^':
            self._eat(); b = self._un(); a = 0 if self.bo else _n(self._sc(a)) ** _n(self._sc(b))
        return a
    def _un(self):
        if self._peek()[1] == '-':
            self._eat(); v = self._un(); return 0 if self.bo else -_n(self._sc(v))
        if self._peek()[1] == '+': self._eat(); return self._un()
        v = self._atom()
        if self._peek()[1] == '%': self._eat(); v = 0 if self.bo else _n(self._sc(v)) / 100
        return v
    def _atom(self):
        k, t = self._peek()
        if k is None: raise KhongTinhDuoc('thiếu toán hạng')
        self._eat()
        if k == 'bool': return t == 'TRUE'
        if k == 'num': return float(t) if ('.' in t or 'e' in t.lower()) else int(t)
        if k == 'str': return t[1:-1].replace('""', '"')
        if k == 'ref': return self._ref(t)
        if k == 'op' and t == '(':
            v = self._cmp()
            if self._eat()[1] != ')': raise KhongTinhDuoc('thiếu ")"')
            return v
        if k == 'fn' and t.upper() in ('IF', 'IFERROR'):
            return self._if(t.upper())
        if k == 'fn':
            args = []
            if self._peek()[1] != ')':
                while True:
                    args.append(self._cmp())
                    if self._peek()[1] == ',': self._eat(); continue
                    break
            if self._eat()[1] != ')': raise KhongTinhDuoc('thiếu ")"')
            return 0 if self.bo else self._ham(t.upper(), args)
        raise KhongTinhDuoc('không hiểu "%s"' % t)
    def _bo(self, bo):
        """Doc 1 doi so; bo=True: chi doc qua, khong tinh (nhanh khong duoc chon cua IF)."""
        cu = self.bo; self.bo = cu or bo
        try: return self._cmp()
        finally: self.bo = cu
    def _if(self, f):
        if f == 'IF':
            c = self._bo(False); cv = False if self.bo else bool(_n(self._sc(c)) if not isinstance(self._sc(c), str) else self._sc(c))
            a = b = False
            if self._peek()[1] == ',': self._eat(); a = self._bo(not cv)
            if self._peek()[1] == ',': self._eat(); b = self._bo(cv)
            if self._eat()[1] != ')': raise KhongTinhDuoc('thiếu ")"')
            return 0 if self.bo else self._sc(a if cv else b)
        i0 = self.i; loi = False
        try: a = self._bo(False)
        except KhongTinhDuoc:
            loi = True; self.i = i0; self._bo(True)
        if self._eat()[1] != ',': raise KhongTinhDuoc('IFERROR thiếu đối số')
        b = self._bo(not loi)
        if self._eat()[1] != ')': raise KhongTinhDuoc('thiếu ")"')
        return 0 if self.bo else self._sc(b if loi else a)
    def tinh_het(self):
        """Tinh truoc moi o cong thuc theo thu tu dong (chuoi tham chieu dai khong bi qua do sau de quy)."""
        o = sorted((ws.title, r, c) for ws in self.wb.worksheets for (r, c), cell in list(ws._cells.items())   # chi o co du lieu
                   if isinstance(cell.value, str) and cell.value.startswith('='))
        thu = {n: i for i, n in enumerate(self.wb.sheetnames)}; o.sort(key=lambda k: (thu[k[0]], k[1], k[2]))
        self.luot = True
        for _ in range(200):                     # moi luot tinh duoc them 1 doan chuoi (do sau <= 120)
            con = []
            for k in o:
                if k in self.cache: continue
                try: self.gia_tri(*k)
                except _QuaSau: self.dang.clear(); con.append(k)
            if not con or len(con) == len(o): break
            o = con
        self.luot = False
    def _ref(self, t):
        if self.bo: return 0
        sh = self.sh
        if '!' in t:
            p, t = t.rsplit('!', 1); sh = p[1:-1].replace("''", "'") if p.startswith("'") else p
            if sh not in self.wb.sheetnames: raise KhongTinhDuoc('không có sheet "%s"' % sh)
        t = t.replace('$', '')
        if re.fullmatch(r'[A-Z]+:[A-Z]+', t):                      # ca cot A:A
            a, b = t.split(':'); t = '%s1:%s%d' % (a, b, max(1, self.wb[sh].max_row))
        if ':' in t:
            a, b = t.split(':')
            ma = re.match(r'([A-Z]+)(\d+)', a); mb = re.match(r'([A-Z]+)(\d+)', b)
            r1, r2 = sorted((int(ma.group(2)), int(mb.group(2)))); c1, c2 = sorted((CI(ma.group(1)), CI(mb.group(1))))
            return ('rng', [self.gia_tri(sh, r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)])
        m = re.match(r'([A-Z]+)(\d+)', t)
        return self.gia_tri(sh, int(m.group(2)), CI(m.group(1)))
    def _sc(self, v):
        if isinstance(v, tuple):
            if len(v[1]) == 1: return v[1][0]
            raise KhongTinhDuoc('dùng vùng ô như 1 giá trị')
        return v
    def _ham(self, f, a):
        def flat():
            for x in a:
                if isinstance(x, tuple):
                    for y in x[1]:
                        if isinstance(y, (int, float)) and not isinstance(y, bool): yield y
                elif isinstance(x, (int, float)): yield x
                elif x not in (None, ''): yield _n(x)
        if f == 'SUM': return sum(flat())
        if f == 'MIN': return min(list(flat()) or [0])
        if f == 'MAX': return max(list(flat()) or [0])
        if f == 'COUNT': return len(list(flat()))
        if f == 'AVERAGE':
            l = list(flat())
            if not l: raise KhongTinhDuoc('AVERAGE rỗng')
            return sum(l) / len(l)
        if f == 'ABS': return abs(_n(self._sc(a[0])))
        if f in ('ROUND', 'ROUNDUP', 'ROUNDDOWN'):
            x = _n(self._sc(a[0])); n = int(_n(self._sc(a[1]))) if len(a) > 1 else 0
            q = Decimal(1).scaleb(-n)
            from decimal import ROUND_UP, ROUND_DOWN
            mode = {'ROUND': ROUND_HALF_UP, 'ROUNDUP': ROUND_UP, 'ROUNDDOWN': ROUND_DOWN}[f]
            return float(Decimal(repr(float(x))).quantize(q, rounding=mode))
        raise KhongTinhDuoc('hàm %s chưa hỗ trợ' % f)

def _n(v):
    if v is None or v == '': return 0
    if isinstance(v, datetime.datetime): d = v - datetime.datetime(1899, 12, 30); return d.days + d.seconds / 86400.0
    if isinstance(v, datetime.date): return (v - datetime.date(1899, 12, 30)).days
    if isinstance(v, bool): return int(v)
    if isinstance(v, (int, float)): return v
    try: return float(str(v).replace(',', '.'))
    except Exception: raise KhongTinhDuoc('"%s" không phải số' % str(v)[:20])
def _txt(v): return '' if v is None else (_so_general(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else str(v))

# =====================================================================================
# 3. FONT TrueType: doc chieu rong chu + cat gon (subset) de nhung vao PDF
# =====================================================================================
class TTF:
    def __init__(self, data, ten):
        self.data = data; self.ten = ten
        n = struct.unpack('>H', data[4:6])[0]; self.tb = {}
        for i in range(n):
            tag, cs, off, ln = struct.unpack('>4sIII', data[12 + 16 * i:28 + 16 * i]); self.tb[tag.decode('latin1')] = (off, ln)
        h = self.tb['head'][0]
        self.upm = struct.unpack('>H', data[h + 18:h + 20])[0]
        self.bbox = struct.unpack('>hhhh', data[h + 36:h + 44]); self.locfmt = struct.unpack('>h', data[h + 50:h + 52])[0]
        hh = self.tb['hhea'][0]
        self.asc, self.desc, self.gap = struct.unpack('>hhh', data[hh + 4:hh + 10]); nhm = struct.unpack('>H', data[hh + 34:hh + 36])[0]
        self.ng = struct.unpack('>H', data[self.tb['maxp'][0] + 4:self.tb['maxp'][0] + 6])[0]
        hm = self.tb['hmtx'][0]
        w = list(struct.unpack('>' + 'Hh' * nhm, data[hm:hm + 4 * nhm]))[0::2]
        self.w = w + [w[-1]] * (self.ng - nhm)
        p = self.tb.get('post', (0, 0))[0]
        self.italic = struct.unpack('>i', data[p + 4:p + 8])[0] / 65536.0 if p else 0
        self.ulpos, self.ulth = struct.unpack('>hh', data[p + 8:p + 12]) if p else (-100, 50)
        o = self.tb.get('OS/2', (0, 0))[0]
        self.cap = struct.unpack('>h', data[o + 88:o + 90])[0] if o and struct.unpack('>H', data[o:o + 2])[0] >= 2 else self.asc
        self.cmap = self._cmap(); self.cw = {}
    def _cmap(self):
        d = self.data; c = self.tb['cmap'][0]; n = struct.unpack('>H', d[c + 2:c + 4])[0]; best = None
        for i in range(n):
            pid, eid, off = struct.unpack('>HHI', d[c + 4 + 8 * i:c + 12 + 8 * i]); fm = struct.unpack('>H', d[c + off:c + off + 2])[0]
            if pid == 3 and eid == 10 and fm == 12: best = (12, c + off); break
            if pid == 3 and eid == 1 and fm == 4 and best is None: best = (4, c + off)
            if pid == 0 and fm == 4 and best is None: best = (4, c + off)
        m = {}
        if best[0] == 4:
            o = best[1]; sc = struct.unpack('>H', d[o + 6:o + 8])[0] // 2
            e = struct.unpack('>%dH' % sc, d[o + 14:o + 14 + 2 * sc]); so = o + 16 + 2 * sc
            s = struct.unpack('>%dH' % sc, d[so:so + 2 * sc]); de = struct.unpack('>%dh' % sc, d[so + 2 * sc:so + 4 * sc])
            ro = so + 4 * sc; r = struct.unpack('>%dH' % sc, d[ro:ro + 2 * sc])
            for k in range(sc):
                for ch in range(s[k], e[k] + 1):
                    if ch == 0xFFFF: continue
                    if r[k] == 0: g = (ch + de[k]) & 0xFFFF
                    else:
                        a = ro + 2 * k + r[k] + 2 * (ch - s[k]); g = struct.unpack('>H', d[a:a + 2])[0]
                        if g: g = (g + de[k]) & 0xFFFF
                    if g: m[ch] = g
        else:
            o = best[1]; ng = struct.unpack('>I', d[o + 12:o + 16])[0]
            for k in range(ng):
                a, b, g = struct.unpack('>III', d[o + 16 + 12 * k:o + 28 + 12 * k])
                for ch in range(a, b + 1): m[ch] = g + ch - a
        return m
    def gid(self, ch): return self.cmap.get(ord(ch), 0)
    def rong(self, s, size):
        cw = self.cw; t = 0
        for ch in s:
            x = cw.get(ch)
            if x is None: x = cw[ch] = self.w[self.cmap.get(ord(ch), 0)]
            t += x
        return t * size / self.upm
    def _glyf(self, g):
        d = self.data; lo = self.tb['loca'][0]
        if self.locfmt == 0: a, b = struct.unpack('>HH', d[lo + 2 * g:lo + 2 * g + 4]); a *= 2; b *= 2
        else: a, b = struct.unpack('>II', d[lo + 4 * g:lo + 4 * g + 8])
        return a, b
    def subset(self, gids):
        """TrueType chi giu cac glyph dung den (giu nguyen so hieu glyph -> CIDToGIDMap Identity)."""
        d = self.data; go = self.tb['glyf'][0]; keep = set(gids) | {0}; st = list(keep)
        while st:                                     # glyph ghep (chu co dau) -> giu ca glyph thanh phan
            g = st.pop(); a, b = self._glyf(g)
            if b - a < 10: continue
            if struct.unpack('>h', d[go + a:go + a + 2])[0] >= 0: continue
            p = go + a + 10
            while True:
                fl, cg = struct.unpack('>HH', d[p:p + 4])
                if cg not in keep: keep.add(cg); st.append(cg)
                p += 4 + (4 if fl & 1 else 2)
                if fl & 8: p += 2
                elif fl & 0x40: p += 4
                elif fl & 0x80: p += 8
                if not fl & 0x20: break
        glyf = bytearray(); loca = []
        for g in range(self.ng):
            loca.append(len(glyf))
            if g in keep:
                a, b = self._glyf(g); glyf += d[go + a:go + b]
                if len(glyf) % 4: glyf += b'\0' * (4 - len(glyf) % 4)
        loca.append(len(glyf))
        tabs = {}
        for t in ('head', 'hhea', 'maxp', 'hmtx', 'cvt ', 'fpgm', 'prep', 'OS/2', 'name', 'post'):
            if t in self.tb: o, l = self.tb[t]; tabs[t] = bytearray(d[o:o + l])
        if 'post' in tabs: tabs['post'] = tabs['post'][:32]; tabs['post'][0:4] = struct.pack('>I', 0x00030000)
        tabs['head'][50:52] = struct.pack('>h', 1); tabs['head'][8:12] = b'\0\0\0\0'
        tabs['loca'] = struct.pack('>%dI' % len(loca), *loca); tabs['glyf'] = bytes(glyf)
        tags = sorted(tabs); n = len(tags); out = bytearray(struct.pack('>IHHHH', 0x00010000, n, 16, 4, n * 16 - 64))
        off = 12 + 16 * n; body = bytearray()
        for t in tags:
            b = bytes(tabs[t]); pad = b + b'\0' * ((4 - len(b) % 4) % 4)
            cs = sum(struct.unpack('>%dI' % (len(pad) // 4), pad)) & 0xFFFFFFFF
            out += struct.pack('>4sIII', t.encode('latin1'), cs, off + len(body), len(b)); body += pad
        return bytes(out + body)

class ThieuFont(Exception): pass

class DoChu:
    """Font chi de DO be rong (khong nhung): doc tu do_rong_chu.py (sinh tu fonts/). Cung so do voi TTF."""
    def __init__(self, k):
        import do_rong_chu as D
        m = D.M.get('%s,%d,%d' % k) or D.M['serif,%d,%d' % (k[1], k[2])]
        self.ten = '%s,%d,%d' % k; self.upm = m['upm']; self.asc = m['asc']; self.desc = m['desc']
        self.ulpos = m['ulpos']; self.ulth = m['ulth']; self.cap = m['cap']; self.w0 = m['w0']
        self.cw = dict(zip(m['chars'], m['w']))
    def rong(self, s, size):
        cw = self.cw; w0 = self.w0
        return sum(cw.get(ch, w0) for ch in s) * size / self.upm

FONT_FILE = {   # (ho, dam, nghieng) -> file
    ('serif', 0, 0): 'LiberationSerif-Regular.ttf', ('serif', 1, 0): 'LiberationSerif-Bold.ttf',
    ('serif', 0, 1): 'LiberationSerif-Italic.ttf', ('serif', 1, 1): 'LiberationSerif-BoldItalic.ttf',
    ('sans', 0, 0): 'Carlito-Regular.ttf', ('sans', 1, 0): 'Carlito-Bold.ttf',
    ('sans', 0, 1): 'Carlito-Italic.ttf', ('sans', 1, 1): 'Carlito-BoldItalic.ttf',
}
DONG = {'serif': 1.3125, 'sans': 1.3636}        # chieu cao dong / co chu (Excel: TNR 12 -> 15.75pt, Calibri 11 -> 15pt)
def ho_font(name):
    n = (name or '').lower()
    return 'sans' if any(k in n for k in ('calibri', 'arial', 'tahoma', 'verdana', 'segoe', 'helvetica', 'carlito')) else 'serif'
# =====================================================================================
# 4. VIET PDF
# =====================================================================================
class PDF:
    def __init__(self, thu_muc_font):
        self.fd = thu_muc_font; self.ttf = {}; self.used = {}; self.pages = []; self.imgs = []; self.vt = []
    def font(self, ho, b, i):
        k = (ho, 1 if b else 0, 1 if i else 0); f = FONT_FILE[k]
        if self.fd is None:                                   # chi do chu (dinh dang in Excel), khong nhung font
            if k not in self.ttf: self.ttf[k] = DoChu(k)
            return self.ttf[k]
        if f not in self.ttf:
            p = os.path.join(self.fd, f)
            if not os.path.exists(p) and os.environ.get('HC_FONT_TAI') and f not in os.environ.get('HC_FONT_BO', '').split(','):
                raise ThieuFont(f)                             # tren web: trang tai font roi chay lai
            if not os.path.exists(p):
                f2 = FONT_FILE[('serif', k[1], k[2])]          # thieu font sans -> dung serif
                if f2 not in self.ttf: self.ttf[f2] = TTF(open(os.path.join(self.fd, f2), 'rb').read(), f2); self.used[f2] = {}
                self.ttf[f] = self.ttf[f2]
            else:
                self.ttf[f] = TTF(open(p, 'rb').read(), f); self.used[f] = {}
        return self.ttf[f]
    def hex(self, f, s):
        u = self.used[f.ten]; o = []
        for ch in s:
            g = f.cmap.get(ord(ch), 0); u[g] = ch; o.append('%04X' % g)
        return '<' + ''.join(o) + '>'
    def anh(self, data):
        """PNG/JPEG -> doi tuong anh PDF. Tra ve chi so."""
        self.imgs.append(_doc_anh(data)); return len(self.imgs) - 1
    def trang(self, w, h, noi_dung, fonts, anhs):
        self.pages.append((w, h, noi_dung, fonts, anhs))
    def luu(self, path):
        objs = []
        def add(b): objs.append(b); return len(objs)
        def stream(dic, data, nen=True):
            if nen: data = zlib.compress(data, 6); dic = dic + b'/Filter/FlateDecode'
            return b'<<' + dic + b'/Length %d>>stream\n' % len(data) + data + b'\nendstream'
        fobj = {}
        names = sorted(set(t.ten for t in self.ttf.values()))
        for i, n in enumerate(names):
            f = [t for t in self.ttf.values() if t.ten == n][0]; u = self.used[n]
            if not u: continue
            ps = 'HC%c%c%c%c%c+' % tuple(chr(65 + (i * 7 + k) % 26) for k in range(5)) + n.split('.')[0].replace('-', '')
            sub = f.subset(list(u)); ff = add(stream(b'/Length1 %d' % len(sub), sub))
            sc = 1000.0 / f.upm
            fl = 2 | (64 if f.italic else 0) | 32
            fdsc = add(('<</Type/FontDescriptor/FontName/%s/Flags %d/FontBBox[%d %d %d %d]/ItalicAngle %d/Ascent %d/Descent %d/CapHeight %d/StemV 80/FontFile2 %d 0 R>>'
                        % (ps, fl, f.bbox[0] * sc, f.bbox[1] * sc, f.bbox[2] * sc, f.bbox[3] * sc, f.italic, f.asc * sc, f.desc * sc, f.cap * sc, ff)).encode())
            gs = sorted(u); W = ' '.join('%d[%d]' % (g, round(f.w[g] * sc)) for g in gs)
            cid = add(('<</Type/Font/Subtype/CIDFontType2/BaseFont/%s/CIDSystemInfo<</Registry(Adobe)/Ordering(Identity)/Supplement 0>>/FontDescriptor %d 0 R/W[%s]/DW 1000/CIDToGIDMap/Identity>>' % (ps, fdsc, W)).encode())
            cm = ['/CIDInit/ProcSet findresource begin 12 dict begin begincmap/CIDSystemInfo<</Registry(Adobe)/Ordering(UCS)/Supplement 0>>def/CMapName/Adobe-Identity-UCS def/CMapType 2 def',
                  '1 begincodespacerange <0000><FFFF> endcodespacerange']
            for k in range(0, len(gs), 100):
                ch = gs[k:k + 100]; cm.append('%d beginbfchar' % len(ch))
                for g in ch:
                    cm.append('<%04X><%s>' % (g, ''.join('%04X' % x for x in _utf16(u[g]))))
                cm.append('endbfchar')
            cm.append('endcmap CMapName currentdict/CMap defineresource pop end end')
            tu = add(stream(b'', '\n'.join(cm).encode()))
            fobj[n] = add(('<</Type/Font/Subtype/Type0/BaseFont/%s/Encoding/Identity-H/DescendantFonts[%d 0 R]/ToUnicode %d 0 R>>' % (ps, cid, tu)).encode())
        iobj = []
        for im in self.imgs:
            w, h, cs, bpc, data, filt, parms, smask = im
            sm = None
            if smask is not None:
                sm = add(stream(b'/Type/XObject/Subtype/Image/Width %d/Height %d/ColorSpace/DeviceGray/BitsPerComponent 8' % (w, h), smask))
            dic = b'/Type/XObject/Subtype/Image/Width %d/Height %d/ColorSpace%s/BitsPerComponent %d' % (w, h, cs, bpc)
            if sm: dic += b'/SMask %d 0 R' % sm
            if filt == 'raw': iobj.append(add(stream(dic, data)))
            else: iobj.append(add(b'<<' + dic + b'/Filter/' + filt.encode() + (parms or b'') + b'/Length %d>>stream\n' % len(data) + data + b'\nendstream'))
        pages_id = len(objs) + 2 * len(self.pages) + 1
        kids = []
        for (w, h, nd, fonts, anhs) in self.pages:
            c = add(stream(b'', nd))
            fr = ''.join('/%s %d 0 R' % (_rname(n), fobj[n]) for n in sorted(fonts) if n in fobj)
            xr = ''.join('/Im%d %d 0 R' % (k, iobj[k]) for k in sorted(anhs))
            kids.append(add(('<</Type/Page/Parent %d 0 R/MediaBox[0 0 %.2f %.2f]/Resources<</Font<<%s>>/XObject<<%s>>>>/Contents %d 0 R>>' % (pages_id, w, h, fr, xr, c)).encode()))
        assert len(objs) + 1 == pages_id
        add(('<</Type/Pages/Kids[%s]/Count %d>>' % (' '.join('%d 0 R' % k for k in kids), len(kids))).encode())
        cat = add(('<</Type/Catalog/Pages %d 0 R>>' % pages_id).encode())
        info = add(('<</Producer(Hoan cong CSKVTT)/CreationDate(D:%s)>>' % time.strftime('%Y%m%d%H%M%S')).encode())
        out = bytearray(b'%PDF-1.5\n%\xe2\xe3\xcf\xd3\n'); xref = []
        for i, o in enumerate(objs):
            xref.append(len(out)); out += b'%d 0 obj\n' % (i + 1) + o + b'\nendobj\n'
        xs = len(out)
        out += b'xref\n0 %d\n0000000000 65535 f \n' % (len(objs) + 1) + b''.join(b'%010d 00000 n \n' % x for x in xref)
        out += b'trailer\n<</Size %d/Root %d 0 R/Info %d 0 R>>\nstartxref\n%d\n%%%%EOF\n' % (len(objs) + 1, cat, info, xs)
        with open(path, 'wb') as fh: fh.write(out)
        return len(out)

def _rname(n): return 'F' + re.sub(r'[^A-Za-z0-9]', '', n.split('.')[0])
def _utf16(ch):
    b = ch.encode('utf-16-be'); return struct.unpack('>%dH' % (len(b) // 2), b)

def _doc_anh(data):
    if data[:2] == b'\xff\xd8':                                     # JPEG: nhung nguyen
        i = 2
        while i < len(data):
            if data[i] != 0xFF: i += 1; continue
            m = data[i + 1]; ln = struct.unpack('>H', data[i + 2:i + 4])[0]
            if m in (0xC0, 0xC1, 0xC2):
                h, w = struct.unpack('>HH', data[i + 5:i + 9]); nc = data[i + 9]
                return (w, h, {1: b'/DeviceGray', 3: b'/DeviceRGB', 4: b'/DeviceCMYK'}[nc], 8, data, 'DCTDecode', None, None)
            i += 2 + ln
        raise ValueError('JPEG lỗi')
    if data[:8] != b'\x89PNG\r\n\x1a\n': raise ValueError('ảnh không phải PNG/JPEG')
    p = 8; idat = b''; plte = b''; trns = b''
    while p < len(data):
        ln, t = struct.unpack('>I4s', data[p:p + 8]); ch = data[p + 8:p + 8 + ln]
        if t == b'IHDR': w, h, bd, ct, _, _, il = struct.unpack('>IIBBBBB', ch)
        elif t == b'PLTE': plte = ch
        elif t == b'tRNS': trns = ch
        elif t == b'IDAT': idat += ch
        elif t == b'IEND': break
        p += 12 + ln
    if il: raise ValueError('PNG xen dòng (interlace) chưa hỗ trợ')
    if bd == 16: raise ValueError('PNG 16-bit chưa hỗ trợ')
    if bd < 8:                                                          # anh chu ky/con dau 1-2-4 bit (gray/bang mau)
        st = (w * bd + 7) // 8; raw = _unfilter(zlib.decompress(idat), st, h, 1, bpp_px=1)
        px = bytearray(); mx = (1 << bd) - 1
        for y in range(h):
            row = raw[y * st:(y + 1) * st]
            for x in range(w):
                b = row[(x * bd) // 8]; sh = 8 - bd - (x * bd) % 8; px.append((b >> sh) & mx)
        if ct == 0:
            g = bytes(int(v * 255 / mx) for v in px)
            return (w, h, b'/DeviceGray', 8, g, 'raw', None, None)
        rgb = bytearray(); a = bytearray()
        for k in px: rgb += plte[3 * k:3 * k + 3]; a.append(trns[k] if k < len(trns) else 255)
        return (w, h, b'/DeviceRGB', 8, bytes(rgb), 'raw', None, bytes(a) if trns else None)
    if ct in (0, 2) and not trns:                                   # khong trong suot: nhung truc tiep
        nc = 1 if ct == 0 else 3
        return (w, h, b'/DeviceGray' if nc == 1 else b'/DeviceRGB', 8, idat, 'FlateDecode', b'/DecodeParms<</Predictor 15/Colors %d/BitsPerComponent 8/Columns %d>>' % (nc, w), None)
    nc = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}[ct]
    raw = _unfilter(zlib.decompress(idat), w, h, nc)
    if ct == 3:
        pal = plte; tr = trns; rgb = bytearray(); a = bytearray()
        for k in raw:
            rgb += pal[3 * k:3 * k + 3]; a.append(tr[k] if k < len(tr) else 255)
        return (w, h, b'/DeviceRGB', 8, bytes(rgb), 'raw', None, bytes(a) if tr else None)
    if ct in (4, 6):
        n = nc - 1; px = bytearray(); a = bytearray()
        mv = memoryview(raw)
        a = bytes(mv[n::nc])
        if n == 3:
            px = bytearray(len(raw) // 4 * 3); px[0::3] = mv[0::4]; px[1::3] = mv[1::4]; px[2::3] = mv[2::4]
        else: px = bytes(mv[0::2])
        return (w, h, b'/DeviceRGB' if n == 3 else b'/DeviceGray', 8, bytes(px), 'raw', None, a)
    return (w, h, b'/DeviceRGB' if nc == 3 else b'/DeviceGray', 8, bytes(raw), 'raw', None, None)

def _unfilter(d, w, h, bpp, bpp_px=None):
    st = w * bpp; bpp = bpp_px or bpp; out = bytearray(st * h); prev = bytearray(st); p = 0
    for y in range(h):
        ft = d[p]; line = bytearray(d[p + 1:p + 1 + st]); p += 1 + st
        if ft == 1:
            for i in range(bpp, st): line[i] = (line[i] + line[i - bpp]) & 255
        elif ft == 2:
            line = bytearray((a + b) & 255 for a, b in zip(line, prev))
        elif ft == 3:
            for i in range(st): line[i] = (line[i] + (((line[i - bpp] if i >= bpp else 0) + prev[i]) >> 1)) & 255
        elif ft == 4:
            for i in range(st):
                a = line[i - bpp] if i >= bpp else 0; b = prev[i]; c = prev[i - bpp] if i >= bpp else 0
                pa = abs(b - c); pb = abs(a - c); pc = abs(a + b - 2 * c)
                line[i] = (line[i] + (a if pa <= pb and pa <= pc else (b if pb <= pc else c))) & 255
        out[y * st:(y + 1) * st] = line; prev = line
    return bytes(out)

# =====================================================================================
# 5. MAU (rgb / theme / indexed)
# =====================================================================================
_THEME_MD = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72']
def _theme(wb):
    t = getattr(wb, 'loaded_theme', None); m = list(_THEME_MD)
    if t:
        try:
            s = t.decode('utf8') if isinstance(t, bytes) else t
            def g(tag):
                x = re.search(r'<a:%s>\s*<a:(?:srgbClr val="([0-9A-Fa-f]{6})"|sysClr[^>]*lastClr="([0-9A-Fa-f]{6})")' % tag, s)
                return (x.group(1) or x.group(2)).upper() if x else None
            for i, tag in enumerate(['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']):
                v = g(tag); m[i] = v or m[i]
        except Exception: pass
    return m
def _tint(hexc, t):
    r, g, b = (int(hexc[i:i + 2], 16) / 255 for i in (0, 2, 4))
    import colorsys
    hh, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 + t) if t < 0 else l * (1 - t) + t
    r, g, b = colorsys.hls_to_rgb(hh, l, s); return (r, g, b)
def mau(c, theme, mac_dinh=None):
    if c is None: return mac_dinh
    try:
        if c.type == 'rgb' and isinstance(c.rgb, str):
            v = c.rgb[-6:]
            if c.rgb.startswith('00') and len(c.rgb) == 8 and v == '000000': return mac_dinh if mac_dinh is not None else (0, 0, 0)
            return tuple(int(v[i:i + 2], 16) / 255 for i in (0, 2, 4)) if not c.tint else _tint(v, c.tint)
        if c.type == 'theme':
            v = theme[c.theme] if c.theme < len(theme) else '000000'
            return _tint(v, c.tint or 0)
        if c.type == 'indexed':
            from openpyxl.styles.colors import COLOR_INDEX
            if c.indexed in (64, 65) or c.indexed >= len(COLOR_INDEX): return mac_dinh
            v = COLOR_INDEX[c.indexed][-6:]; return tuple(int(v[i:i + 2], 16) / 255 for i in (0, 2, 4))
    except Exception: pass
    return mac_dinh

# =====================================================================================
# 6. DAN TRANG + VE 1 SHEET
# =====================================================================================
GIAY = {9: (595.28, 841.89), 8: (841.89, 1190.55), 1: (612, 792), 5: (612, 1008), 11: (419.53, 595.28)}
_VUNG = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?")

def _vung_in(ws):
    pa = ws.print_area
    if pa:
        s = (pa if isinstance(pa, str) else pa[0]).split(',')[0].split('!')[-1]
        m = _VUNG.search(s)
        if m:
            r1, c1 = int(m.group(2)), CI(m.group(1)); r2, c2 = (int(m.group(4)), CI(m.group(3))) if m.group(3) else (r1, c1)
            return r1, c1, r2, c2
    r2 = c2 = 1
    for row in ws.iter_rows():
        for c in row:
            if c.value not in (None, '') and not isinstance(c, MergedCell): r2 = max(r2, c.row); c2 = max(c2, c.column)
    for m in ws.merged_cells.ranges:
        if m.min_row <= r2 and m.min_col <= c2: r2 = max(r2, m.max_row) if m.min_row <= r2 else r2; c2 = max(c2, m.max_col)
    return 1, 1, r2, c2

class Sheet:
    """Hinh hoc 1 sheet: do rong cot, chieu cao dong (pt), o gop, chu hien thi."""
    def __init__(self, ws, ev, pdf, theme, mdw):
        self.ws = ws; self.ev = ev; self.pdf = pdf; self.theme = theme; self.mdw = mdw
        self.r1, self.c1, self.r2, self.c2 = _vung_in(ws)
        self.mg = {}; self.mgcell = {}
        for m in ws.merged_cells.ranges:
            self.mg[(m.min_row, m.min_col)] = (m.max_row, m.max_col)
            for r in range(m.min_row, m.max_row + 1):
                for c in range(m.min_col, m.max_col + 1): self.mgcell[(r, c)] = (m.min_row, m.min_col)
        dcw = ws.sheet_format.defaultColWidth
        base = ws.sheet_format.baseColWidth or 8
        dw = dcw if dcw else math.trunc((base * mdw + 5) / mdw * 256) / 256
        W = {}
        for k, d in ws.column_dimensions.items():
            lo = d.min or CI(k); hi = d.max or lo
            for c in range(lo, hi + 1):
                W[c] = 0 if d.hidden else (d.width if d.width else dw)
        self._W = W; self._dw = dw
        self.cw = {c: self._px(W.get(c, dw)) * 0.75 for c in range(self.c1, self.c2 + 1)}
        self.rows = [r for r in range(self.r1, self.r2 + 1) if not ws.row_dimensions[r].hidden]
        self.cols = [c for c in range(self.c1, self.c2 + 1) if self.cw[c] > 0]
        self.rh = {}
        dfh = ws.sheet_format.defaultRowHeight or 15
        for r in self.rows:
            h = ws.row_dimensions[r].ht
            self.rh[r] = h if h is not None else self._tu_cao(r, dfh)
    def _px(self, w):
        m = self.mdw
        return math.trunc(((256 * w + math.trunc(128 / m)) / 256) * m)
    # ---- noi dung o
    def o(self, r, c):
        return self.ws.cell(r, c)
    def chu(self, cell):
        """-> (runs, la_so). runs = [(text, b, i, u, strike, sz, ho, mau)]"""
        v = cell.value
        f = cell.font; base = (bool(f.b), bool(f.i), f.u, bool(f.strike), float(f.sz or 11), ho_font(f.name), mau(f.color, self.theme, (0, 0, 0)))
        if v is None: return [], False
        if isinstance(v, str): v = unicodedata.normalize('NFC', v)        # dau roi (to hop) -> chu dung san
        if type(v).__name__ == 'CellRichText':
            runs = []
            for t in v:
                if isinstance(t, str): runs.append((unicodedata.normalize('NFC', t),) + base)
                else:
                    g = t.font
                    runs.append((unicodedata.normalize('NFC', t.text), bool(g.b) if g.b is not None else base[0], bool(g.i) if g.i is not None else base[1],
                                 g.u if g.u is not None else base[2], bool(g.strike) if g.strike is not None else base[3],
                                 float(g.sz) if g.sz else base[4], ho_font(g.rFont) if g.rFont else base[5], mau(g.color, self.theme, base[6])))
            return runs, False
        if isinstance(v, str) and v.startswith('='):
            v = self.ev.gia_tri(self.ws.title, cell.row, cell.column)
            if v is None: return [], False
        so = isinstance(v, (int, float)) and not isinstance(v, bool)
        return [(dinh_dang(v, cell.number_format),) + base], so
    def rong_cot(self, c):
        if c in self.cw: return self.cw[c]
        return self._px(self._W.get(c, self._dw)) * 0.75
    def rong_o(self, r, c):
        if (r, c) in self.mg:
            r2, c2 = self.mg[(r, c)]; return sum(self.rong_cot(x) for x in range(c, c2 + 1))
        return self.rong_cot(c)
    def _tu_cao(self, r, dfh):
        """Chieu cao tu dong cho dong khong dat chieu cao (nhu Excel/LibreOffice)."""
        best = 0
        for c in range(self.c1, self.c2 + 1):
            if (r, c) in self.mgcell:
                a = self.mgcell[(r, c)]
                if a != (r, c) or self.mg[a][0] > r: continue
            cell = self.ws.cell(r, c)
            if cell.value is None: continue
            runs, so = self.chu(cell)
            if not runs: continue
            if cell.alignment.wrap_text and not so:
                lines = self.ngat(runs, max(1, self.rong_o(r, c) - 3 - self._lui(cell)))
                h = sum(l[1] for l in lines)
            else:
                h = max(rn[5] * DONG[rn[6]] for rn in runs)
            best = max(best, h)
        f = self.ws.parent._fonts[0] if self.ws.parent._fonts else None
        mn = (float(f.sz) * DONG[ho_font(f.name)]) if f is not None and f.sz else dfh
        return max(best, mn if best == 0 else 0, best)
    def _lui(self, cell):
        return (cell.alignment.indent or 0) * self.mdw * 3 * 0.75
    # ---- ngat dong
    def ngat(self, runs, wmax):
        """runs -> [(cac doan [(text, run)], cao_dong, rong)]"""
        pdf = self.pdf; lines = []; cur = []; cw = 0.0; lh = 0.0
        def flush():
            nonlocal cur, cw, lh
            while cur and cur[-1][0].isspace(): cw -= cur[-1][2]; cur.pop()
            lines.append((cur, lh or self._lh_mac(runs), cw)); cur = []; cw = 0.0; lh = 0.0
        for rn in runs:
            f = pdf.font(rn[6], rn[1], rn[2]); sz = rn[5]; h = sz * DONG[rn[6]]
            parts = rn[0].split('\n')
            for pi, part in enumerate(parts):
                if pi > 0: lh = max(lh, h); flush()
                for tok in re.findall(r'\s+|\S+', part):
                    w = f.rong(tok, sz)
                    if tok.isspace():
                        if cur: cur.append((tok, rn, w)); cw += w; lh = max(lh, h)
                        continue
                    if cw + w > wmax + 0.01 and cur: flush()
                    while w > wmax + 0.01 and len(tok) > 1:
                        cs = [0.0]
                        for ch in tok: cs.append(cs[-1] + f.rong(ch, sz))          # cong don be rong tung chu
                        lo_, hi_ = 1, len(tok)
                        while lo_ < hi_:                                          # k lon nhat ma tok[:k] vua phan con lai
                            mid = (lo_ + hi_ + 1) // 2
                            if cs[mid] <= wmax - cw + 0.01: lo_ = mid
                            else: hi_ = mid - 1
                        k = lo_
                        if cur and k < len(tok) and cs[k] < wmax * 0.3: flush(); continue
                        cur.append((tok[:k], rn, cs[k])); lh = max(lh, h); cw += cs[k]; flush()
                        tok = tok[k:]; w = cs[-1] - cs[k]
                    cur.append((tok, rn, w)); cw += w; lh = max(lh, h)
        if cur or not lines: lh = lh or max(r[5] * DONG[r[6]] for r in runs); flush()
        return lines
    def _lh_mac(self, runs):
        return max(r[5] * DONG[r[6]] for r in runs) if runs else 15

def _hf(text):
    """Chuoi header/footer Excel -> danh sach doan (chu, dam, nghieng, co)."""
    return text

class Ve:
    """Ghi lenh ve PDF cho 1 trang (toa do trang, goc duoi trai)."""
    def __init__(self, pdf, H):
        self.pdf = pdf; self.H = H; self.o = []; self.fonts = set(); self.anhs = set(); self.vt = []
    def mau_to(self, c): self.o.append('%.3f %.3f %.3f rg' % c)
    def mau_net(self, c): self.o.append('%.3f %.3f %.3f RG' % c)
    def hcn(self, x, y, w, h, c):
        self.mau_to(c); self.o.append('%.2f %.2f %.2f %.2f re f' % (x, self.H - y - h, w, h))
    def net(self, x1, y1, x2, y2, w, kieu=None, c=(0, 0, 0)):
        self.mau_net(c); d = {'dashed': '[3 2] 0 d', 'dotted': '[1 1] 0 d', 'hair': '[0.5 0.5] 0 d', 'mediumDashed': '[4 2] 0 d',
                             'dashDot': '[3 1 1 1] 0 d', 'mediumDashDot': '[4 1 1 1] 0 d', 'dashDotDot': '[3 1 1 1 1 1] 0 d',
                             'mediumDashDotDot': '[4 1 1 1 1 1] 0 d', 'slantDashDot': '[3 1 1 1] 0 d'}.get(kieu, '[] 0 d')
        self.o.append('%s %.2f w %.2f %.2f m %.2f %.2f l S' % (d, w, x1, self.H - y1, x2, self.H - y2))
    def chu(self, x, y, s, ten_font, f, sz, c):
        if not s: return
        self.fonts.add(ten_font); self.mau_to(c)
        self.o.append('BT /%s %.2f Tf %.2f %.2f Td %s Tj ET' % (_rname(ten_font), sz, x, self.H - y, self.pdf.hex(f, s)))
    def cat(self, x, y, w, h): self.o.append('q %.2f %.2f %.2f %.2f re W n' % (x, self.H - y - h, w, h))
    def bo_cat(self): self.o.append('Q')
    def anh(self, k, x, y, w, h):
        self.anhs.add(k); self.o.append('q %.2f 0 0 %.2f %.2f %.2f cm /Im%d Do Q' % (w, h, x, self.H - y - h, k))
    def xong(self): return '\n'.join(self.o).encode('latin1')

NET = {'hair': 0.25, 'thin': 0.5, 'medium': 1.0, 'thick': 1.5, 'double': 0.5, 'dashed': 0.5, 'dotted': 0.5,
       'mediumDashed': 1.0, 'dashDot': 0.5, 'mediumDashDot': 1.0, 'dashDotDot': 0.5, 'mediumDashDotDot': 1.0, 'slantDashDot': 1.0}

def _phan_trang(S, avail_h, tieu_de):
    """Chia dong thanh trang -> [[dong...]] (chua gom dong tieu de lap; tieu de lap o trang co dong dau > dong tieu de cuoi)."""
    ws = S.ws; brk = set(b.id for b in ws.row_breaks.brk if b.id)
    th = sum(S.rh[r] for r in tieu_de if r in S.rh); last = max(tieu_de) if tieu_de else 0
    pages = []; cur = []; h = 0.0
    for r in S.rows:
        rh = S.rh[r]
        if cur and h + rh > avail_h + 0.01:
            pages.append(cur); cur = []
        if not cur: h = th if (tieu_de and pages and r > last) else 0.0
        cur.append(r); h += rh
        if r in brk: pages.append(cur); cur = []
    if cur: pages.append(cur)
    return pages or [[]]

def _tieu_de(ws):
    t = ws.print_title_rows
    if not t: return []
    m = re.search(r'\$?(\d+):\$?(\d+)', t)
    return list(range(int(m.group(1)), int(m.group(2)) + 1)) if m else []

def kho_giay(ws, S):
    """-> (pw, ph, ml, mr, mt, mb, aw, ah, k): kho giay, le, vung in duoc, ti le co (nhu Excel)."""
    ps = ws.page_setup; pm = ws.page_margins
    pw, ph = GIAY.get(int(ps.paperSize or 9), GIAY[9])
    if ps.orientation == 'landscape': pw, ph = ph, pw
    ml, mr, mt, mb = pm.left * 72, pm.right * 72, pm.top * 72, pm.bottom * 72
    aw, ah = pw - ml - mr, ph - mt - mb
    tw = sum(S.cw[c] for c in S.cols) or 1
    th_all = sum(S.rh[r] for r in S.rows) or 1
    spr = ws.sheet_properties.pageSetUpPr
    if spr is not None and spr.fitToPage:
        fw = ps.fitToWidth if ps.fitToWidth is not None else 1; fh = ps.fitToHeight if ps.fitToHeight is not None else 1
        k = 1.0
        if fw: k = min(k, aw * fw / tw)
        if fh: k = min(k, ah * fh / th_all)
        k = max(0.1, math.floor(k * 100 + 1e-9) / 100)
    else:
        k = (ps.scale or 100) / 100.0
    return pw, ph, ml, mr, mt, mb, aw, ah, k

def chia_trang(ws, S):
    """-> (trang [[dong]], tieu_de, k) giong luc in."""
    k = kho_giay(ws, S)[8]
    tieu_de = [r for r in _tieu_de(ws) if r in S.rh]
    return _phan_trang(S, kho_giay(ws, S)[7] / k, tieu_de), tieu_de, k

def giu_khoi_ky(ws, S):
    """Khoi ky (tu dong 'ĐẠI DIỆN BÊN …' / 'Chỉ huy trưởng' cuoi bien ban den het vung in) phai nam tron 1 trang:
    neu bi chia cat -> chen ngat trang ngay truoc khoi. Tra ve dong bat dau khoi neu da chen."""
    from openpyxl.worksheet.pagebreak import Break
    rows = S.rows
    if not rows: return None
    bd = None
    for r in reversed(rows[-40:]):
        for c in S.cols:
            v = ws.cell(r, c).value
            if isinstance(v, str) and re.match(r'\s*(ĐẠI DIỆN BÊN|Chỉ huy trưởng|CHỈ HUY TRƯỞNG|NGƯỜI LẬP|Người lập)', v):
                bd = r; break
        if bd == r: break
    if bd is None: return None
    ids = [b.id for b in ws.row_breaks.brk if b.id and bd <= b.id < rows[-1]]
    if ids: ws.row_breaks.brk = [b for b in ws.row_breaks.brk if b.id not in ids]
    trang, tieu_de, k = chia_trang(ws, S)
    khoi = [r for r in rows if r >= bd]
    o = [i for i, p in enumerate(trang) if any(r in p for r in khoi)]
    if len(o) <= 1: return None
    ah = kho_giay(ws, S)[7] / k
    cao = sum(S.rh[r] for r in khoi) + (sum(S.rh[r] for r in tieu_de) if tieu_de and bd > max(tieu_de) else 0)
    if cao > ah: return None                       # khoi ky dai hon 1 trang: khong the giu tron
    keo = bd
    if tieu_de and bd > max(tieu_de):              # trang ky co tieu de bang lap -> keo theo toi da 3 dong cuoi bang
        truoc_ = [r for r in rows if max(tieu_de) < r < bd and S.rh[r] > 0]
        while truoc_ and len([r for r in rows if keo <= r < bd]) < 3:
            r0 = truoc_.pop()
            if cao + S.rh[r0] > ah: break
            cao += S.rh[r0]; keo = r0
    truoc = max([r for r in rows if r < keo] or [0])
    if truoc: ws.row_breaks.append(Break(id=truoc))
    return bd

def giu_tieu_de(ws, S):
    """Dong tieu de bang (lap lai moi trang) khong duoc nam cuoi trang ma khong co dong du lieu nao di kem."""
    from openpyxl.worksheet.pagebreak import Break
    trang, tieu_de, k = chia_trang(ws, S)
    if not tieu_de: return None
    for i, p in enumerate(trang[:-1]):
        sau = [r for r in p if r > max(tieu_de)]
        if any(r in tieu_de for r in p) and not sau and p[0] < min(tieu_de):
            truoc = max([r for r in S.rows if r < min(tieu_de)] or [0])
            if truoc and truoc not in [b.id for b in ws.row_breaks.brk]:
                ws.row_breaks.append(Break(id=truoc)); return min(tieu_de)
    return None

def rong_tran(S, r, c, ha):
    """Be rong co the dung cho chu KHONG xuong dong cua o (r,c): ca o tran sang o trong ben canh (nhu Excel)."""
    w = S.rong_o(r, c)
    if (r, c) in S.mg: return w
    cols = S.cols
    if c not in cols: return w
    i = cols.index(c)
    def trong(cc): return (r, cc) not in S.mgcell and S.ws.cell(r, cc).value in (None, '')
    R_ = L_ = 0.0; j = i
    while j + 1 < len(cols) and trong(cols[j + 1]): j += 1; R_ += S.cw[cols[j]]
    j = i
    while j - 1 >= 0 and trong(cols[j - 1]): j -= 1; L_ += S.cw[cols[j]]
    if ha in ('left', 'general', 'justify', 'fill', None): return w + R_
    if ha == 'right': return w + L_
    return w + 2 * min(L_, R_)

def ve_sheet(ws, ev, pdf, theme, mdw, canh_bao):
    S = Sheet(ws, ev, pdf, theme, mdw); S.canh = canh_bao
    pw, ph, ml, mr, mt, mb, aw, ah, k = kho_giay(ws, S)
    trang, tieu_de, k = chia_trang(ws, S)
    def co_noi_dung(rows):                       # Excel/LibreOffice khong in trang trong
        for r in rows:
            for c in S.cols:
                a = S.mgcell.get((r, c), (r, c)); cell = ws.cell(*a)
                if cell.value not in (None, ''): return True
                b = ws.cell(r, c).border
                if b is not None and any(getattr(b, sd) is not None and getattr(b, sd).style for sd in ('left', 'right', 'top', 'bottom')): return True
                if cell.fill is not None and cell.fill.fill_type == 'solid': return True
        return False
    trang = [p for p in trang if co_noi_dung(p)]
    if not trang: return 0                                     # sheet trong: Excel khong in
    # chia cot neu qua rong (khi khong vua 1 trang ngang)
    nhom_cot = []; cur = []; w = 0
    for c in S.cols:
        if cur and w + S.cw[c] > aw / k + 0.01: nhom_cot.append(cur); cur = []; w = 0
        cur.append(c); w += S.cw[c]
    nhom_cot.append(cur)
    hc = bool(ws.print_options.horizontalCentered); vc = bool(ws.print_options.verticalCentered)
    tong = len(trang) * len(nhom_cot); so = 0
    anh = []
    for im in getattr(ws, '_images', []):
        try: anh.append((im, pdf.anh(im._data())))
        except Exception as e: canh_bao.append('%s: bỏ qua 1 ảnh (%s)' % (ws.title, e))
    for cols in nhom_cot:
        for rows in trang:
            so += 1
            last_title = max(tieu_de) if tieu_de else 0
            rr = ([r for r in tieu_de] if (tieu_de and rows and rows[0] > last_title) else []) + rows
            V = Ve(pdf, ph)
            cw_ = sum(S.cw[c] for c in cols) * k; chh = sum(S.rh[r] for r in rr) * k
            ox = ml + ((aw - cw_) / 2 if hc else 0); oy = mt + ((ah - chh) / 2 if vc else 0)
            X = {}; x = 0.0
            for c in cols: X[c] = x; x += S.cw[c]
            Y = {}; y = 0.0
            for r in rr: Y[r] = y; y += S.rh[r]
            _ve_trang(S, V, rr, cols, X, Y, ox, oy, k)
            for im, idx in anh:
                _ve_anh(S, V, im, idx, rr, cols, X, Y, ox, oy, k)
            _ve_hf(ws, V, pdf, pw, ph, ml, mr, so, tong, k)
            pdf.trang(pw, ph, V.xong(), V.fonts, V.anhs); pdf.vt.append((ws.title, V.vt))
    return tong

def _ve_anh(S, V, im, idx, rr, cols, X, Y, ox, oy, k):
    a = im.anchor
    try:
        fr = a._from; c0, r0 = fr.col + 1, fr.row + 1
        if r0 not in Y or c0 not in X: return
        x = X[c0] + fr.colOff / 12700.0; y = Y[r0] + fr.rowOff / 12700.0
        if hasattr(a, 'to') and a.to is not None:
            to = a.to; c1, r1 = to.col + 1, to.row + 1
            x2 = (X[c1] if c1 in X else X[c0] + sum(S.cw.get(c, 0) for c in range(c0, c1))) + to.colOff / 12700.0
            y2 = (Y[r1] if r1 in Y else Y[r0] + sum(S.rh.get(r, 0) for r in range(r0, r1))) + to.rowOff / 12700.0
        else:
            x2 = x + a.ext.width / 12700.0; y2 = y + a.ext.height / 12700.0
        V.anh(idx, ox + x * k, oy + y * k, (x2 - x) * k, (y2 - y) * k)
    except Exception:
        pass

def _hf_parse(t):
    """'&"Times New Roman,Italic"&11Trang &P / &N' -> (chu, font, kieu, co)"""
    fn, st, sz = None, '', None
    m = re.search(r'&"([^",]*),?([^"]*)"', t)
    if m: fn, st = m.group(1), m.group(2); t = t.replace(m.group(0), '')
    m = re.search(r'&(\d+)', t)
    if m: sz = float(m.group(1)); t = t.replace(m.group(0), '', 1)
    return t, fn, st, sz

def _ve_hf(ws, V, pdf, pw, ph, ml, mr, so, tong, k):
    pm = ws.page_margins
    hfo = getattr(ws, 'HeaderFooter', None)
    dau = bool(hfo is not None and hfo.differentFirst) and so == 1
    chan = bool(hfo is not None and hfo.differentOddEven) and so % 2 == 0
    for hf, top in (((ws.firstHeader if dau else ws.evenHeader if chan else ws.oddHeader), True),
                    ((ws.firstFooter if dau else ws.evenFooter if chan else ws.oddFooter), False)):
        for part, al in ((hf.left, 'l'), (hf.center, 'c'), (hf.right, 'r')):
            if not part or not part.text: continue
            t, fn, st, sz = _hf_parse(part.text)
            fn = part.font.split(',')[0] if part.font else (fn or 'Times New Roman')
            st = (part.font.split(',')[1] if part.font and ',' in part.font else st or '').lower()
            sz = float(part.size or sz or 11)
            t = t.replace('&P', str(so)).replace('&N', str(tong)).replace('&A', ws.title).replace('&D', time.strftime('%d/%m/%Y')).replace('&T', time.strftime('%H:%M')).replace('&&', '&')
            t = re.sub(r'&[A-Za-z]', '', t)
            f = pdf.font(ho_font(fn), 'bold' in st, 'italic' in st); sz2 = sz * k
            w = f.rong(t, sz2)
            x = ml if al == 'l' else (pw - mr - w if al == 'r' else (pw - w) / 2)
            y = (pm.header * 72 + f.asc / f.upm * sz2) if top else (ph - pm.footer * 72 - (-f.desc) / f.upm * sz2)
            V.chu(x, y, t, f.ten, f, sz2, (0, 0, 0))

def _ve_trang(S, V, rr, cols, X, Y, ox, oy, k):
    ws = S.ws; theme = S.theme; colset = set(cols); rowset = set(rr)
    def R(r, c):                       # hinh chu nhat o (r,c) hoac o gop (toa do sheet)
        if (r, c) in S.mg:
            r2, c2 = S.mg[(r, c)]
            w = sum(S.cw[x] for x in range(c, c2 + 1) if x in colset); h = sum(S.rh[y] for y in range(r, r2 + 1) if y in rowset)
            return X[c], Y[r], w, h
        return X[c], Y[r], S.cw[c], S.rh[r]
    P = lambda x, y: (ox + x * k, oy + y * k)
    # --- lop 1: nen
    for r in rr:
        for c in cols:
            a = S.mgcell.get((r, c))
            if a and a != (r, c): continue
            cell = ws.cell(r, c); fl = cell.fill
            if fl is not None and fl.fill_type == 'solid':
                col = mau(fl.fgColor, theme, None)
                if col and col != (1, 1, 1):
                    x, y, w, h = R(r, c); px, py = P(x, y); V.hcn(px, py, w * k, h * k, col)
    # --- lop 2: chu
    for r in rr:
        for c in cols:
            a = S.mgcell.get((r, c))
            if a and a != (r, c):
                if not (a[0] not in rowset and r == min(x for x in rr if S.mgcell.get((x, c)) == a) and c == a[1]) and \
                   not (a[1] not in colset and c == min(x for x in cols if S.mgcell.get((r, x)) == a) and r == a[0]):
                    continue
                cell = ws.cell(*a); x0 = X[c]; y0 = Y[r]
                r2, c2 = S.mg[a]
                w = sum(S.cw[x] for x in range(c, c2 + 1) if x in colset); h = sum(S.rh[y] for y in range(r, r2 + 1) if y in rowset)
            else:
                cell = ws.cell(r, c)
                if cell.value is None: continue
                x0, y0, w, h = R(r, c)
            _ve_o(S, V, cell, r, c, x0, y0, w, h, cols, colset, X, P, k)
    # --- lop 3: vien
    for r in rr:
        for c in cols:
            b = ws.cell(r, c).border
            if b is None: continue
            x, y, w, h = X[c], Y[r], S.cw[c], S.rh[r]
            a0 = S.mgcell.get((r, c))
            for side, (x1, y1, x2, y2), nb in (('top', (x, y, x + w, y), (r - 1, c)), ('bottom', (x, y + h, x + w, y + h), (r + 1, c)),
                                               ('left', (x, y, x, y + h), (r, c - 1)), ('right', (x + w, y, x + w, y + h), (r, c + 1))):
                sd = getattr(b, side)
                if sd is None or not sd.style: continue
                if a0 is not None and S.mgcell.get(nb) == a0: continue          # vien BEN TRONG o gop: Excel khong ve
                col = mau(sd.color, theme, (0, 0, 0)) or (0, 0, 0)
                a1, b1 = P(x1, y1); a2, b2 = P(x2, y2); lw = NET.get(sd.style, 0.5)
                if sd.style == 'double':
                    d = 0.75
                    if y1 == y2: V.net(a1, b1 - d, a2, b2 - d, 0.4, None, col); V.net(a1, b1 + d, a2, b2 + d, 0.4, None, col)
                    else: V.net(a1 - d, b1, a2 - d, b2, 0.4, None, col); V.net(a1 + d, b1, a2 + d, b2, 0.4, None, col)
                else:
                    V.net(a1, b1, a2, b2, lw, sd.style, col)

PAD = 1.5   # le trai/phai trong o (pt) - hieu chinh: khop 99,4% voi 1.809 dong Excel tu canh chieu cao (file mau)

def _ve_o(S, V, cell, r, c, x0, y0, w, h, cols, colset, X, P, k):
    runs, so = S.chu(cell)
    if not runs or not ''.join(rn[0] for rn in runs).strip(): return
    if h < 10 and r not in S.mg and max(rn[5] for rn in runs) * DONG[runs[0][6]] > h + 1: return   # dong dem cua mau: chu khong lot -> khong ve vet
    al = cell.alignment; ha = al.horizontal or 'general'; va = al.vertical or 'bottom'
    if ha == 'general': ha = 'right' if so else 'left'
    if ha in ('centerContinuous', 'distributed'): ha = 'center' if ha == 'centerContinuous' else 'justify'
    if ha == 'fill': ha = 'left'
    lui = S._lui(cell)
    wrap = bool(al.wrap_text) and not so
    xx, ww = x0, w
    if wrap or ha == 'justify':
        lines = S.ngat(runs, max(1.0, w - 2 * PAD - lui))
    else:
        lines = S.ngat(runs, 1e9)
        lines = [lines[0]] if lines else lines
        tw = lines[0][2] if lines else 0
        if so and tw > w - 2 * PAD + 0.01:              # so khong vua -> ####
            f = S.pdf.font(runs[0][6], runs[0][1], runs[0][2]); n = max(1, int((w - 2 * PAD) / max(0.1, f.rong('#', runs[0][5]))))
            rn = runs[0]; lines = [([('#' * n, rn, f.rong('#' * n, rn[5]))], rn[5] * DONG[rn[6]], f.rong('#' * n, rn[5]))]
        elif tw > w - 2 * PAD and (r, c) not in S.mg:   # chu tran sang o trong ben canh
            i = cols.index(c)
            def trong(cc):
                if (r, cc) in S.mgcell: return False
                v = S.ws.cell(r, cc).value
                return v is None or v == ''
            if ha == 'left':
                j = i
                while xx + ww - x0 < tw + 2 * PAD + lui and j + 1 < len(cols) and trong(cols[j + 1]): j += 1; ww += S.cw[cols[j]]
            elif ha == 'right':
                j = i
                while ww < tw + 2 * PAD and j - 1 >= 0 and trong(cols[j - 1]): j -= 1; ww += S.cw[cols[j]]; xx -= S.cw[cols[j]]
            elif ha == 'center':
                L_ = R_ = i
                while ww < tw + 2 * PAD:
                    moved = False
                    if R_ + 1 < len(cols) and trong(cols[R_ + 1]) and (L_ - 1 < 0 or not trong(cols[L_ - 1]) or True):
                        R_ += 1; ww += S.cw[cols[R_]]; moved = True
                    if L_ - 1 >= 0 and trong(cols[L_ - 1]) and ww < tw + 2 * PAD:
                        L_ -= 1; ww += S.cw[cols[L_]]; xx -= S.cw[cols[L_]]; moved = True
                    if not moved: break
                # can giua theo o goc: lay doi xung quanh o goc
                ctr = x0 + w / 2; half = min(ctr - xx, xx + ww - ctr)
                if half * 2 >= tw: xx, ww = ctr - half, 2 * half                 # doi xung quanh o goc (nhu Excel)
    tot = sum(l[1] for l in lines)
    canh = getattr(S, 'canh', None)
    if canh is not None and S.rh.get(r, 99) >= 10:
        o = '%s!%s%d' % (S.ws.title, CL(c), r)
        if not (wrap or ha == 'justify') and lines and lines[0][2] > ww - lui + 0.5 and not so:
            canh.append('%s: chữ dài hơn ô/vùng in, phần cuối bị cắt khi in ("…%s")' % (o, ''.join(rn[0] for rn in runs).strip()[-25:]))
        elif len(lines) > 1 and tot - lines[-1][1] * 0.5 > h:          # mat it nhat nua dong cuoi
            canh.append('%s: chữ nhiều dòng hơn chiều cao dòng, bị cắt khi in' % o)
    if va == 'top': yy = y0 + 0.75
    elif va in ('center', 'justify', 'distributed'): yy = y0 + (h - tot) / 2
    else: yy = y0 + h - tot - 0.75
    px, py = P(xx, y0); V.cat(px, py, ww * k, h * k); V.vt.append((r, c, px, py, ww * k, h * k))
    for li, (segs, lh, lw) in enumerate(lines):
        if not segs: yy += lh; continue
        avail = ww - 2 * PAD - lui
        if ha == 'right': x = xx + ww - PAD - lw
        elif ha == 'center': x = xx + (ww - lw) / 2
        else: x = xx + PAD + lui
        gap = 0.0
        if ha == 'justify' and li < len(lines) - 1 and wrap:
            ns = sum(1 for s in segs if s[0].isspace())
            if ns: gap = max(0.0, (avail - lw) / ns)
        mx = max(s[1][5] for s in segs)
        f0 = S.pdf.font(segs[0][1][6], segs[0][1][1], segs[0][1][2])
        base = yy + (lh - mx * (f0.asc - f0.desc) / f0.upm) / 2 + mx * f0.asc / f0.upm
        # vi tri tung doan (dong gian deu: cong khoang sau moi dau cach), roi gom doan LIEN NHAU cung kieu -> 1 chuoi
        groups = []; xc = x
        for t, rn, sw in segs:
            if groups and groups[-1][1] is rn and abs(groups[-1][3] + groups[-1][2] - xc) < 0.01:
                groups[-1][0] += t; groups[-1][2] += sw
            else:
                groups.append([t, rn, sw, xc])
            xc += sw + (gap if t.isspace() else 0)
        for t, rn, sw, x in groups:
            f = S.pdf.font(rn[6], rn[1], rn[2]); col = rn[7] or (0, 0, 0)
            if t:
                a1, b1 = P(x, base); V.chu(a1, b1, t, f.ten, f, rn[5] * k, col)
            if rn[3] and rn[3] != 'none' and sw > 0:
                uy = base - f.ulpos / f.upm * rn[5]; a2, b2 = P(x, uy); lw = max(0.3, f.ulth / f.upm * rn[5]) * k
                V.net(a2, b2, a2 + sw * k, b2, lw, None, col)
                if rn[3] in ('double', 'doubleAccounting'):
                    a3, b3 = P(x, uy + rn[5] * 0.12); V.net(a3, b3, a3 + sw * k, b3, lw, None, col)
            if rn[4] and t.strip():
                sy = base - f.cap / f.upm * rn[5] * 0.45; a2, b2 = P(x, sy)
                V.net(a2, b2, a2 + sw * k, b2, max(0.3, f.ulth / f.upm * rn[5]) * k, None, col)
        yy += lh
    V.bo_cat()

# =====================================================================================
# 7. HAM CHINH
# =====================================================================================
def _mdw(wb, pdf):
    f = wb._fonts[0] if wb._fonts else None
    ho = ho_font(f.name if f is not None else 'Calibri'); sz = float(f.sz) if f is not None and f.sz else 11
    t = pdf.font(ho, 0, 0); px = sz * 96 / 72
    return max(1, round(max(t.rong(d, px) for d in '0123456789')))

def cac_o_in(ws, ev):
    """Cac o se in ra (trong vung in, dong/cot khong an): (dong, cot, chu hien thi). Dung cho kiem tra."""
    pdf = PDF(os.environ.get('HC_FONTS', os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'fonts')))
    S = Sheet(ws, ev, pdf, _theme(ws.parent), _mdw(ws.parent, pdf))
    for r in S.rows:
        for c in S.cols:
            a = S.mgcell.get((r, c))
            if a and a != (r, c): continue
            cell = ws.cell(r, c)
            if cell.value is None: continue
            runs, so = S.chu(cell)
            t = ''.join(rn[0] for rn in runs)
            if t.strip(): yield r, c, t

def xuat(src, out, thu_muc_font, log=print):
    """Xuat toan bo sheet dang hien cua file src -> out (PDF). Tra ve {'trang': [(sheet, so trang)], 'canh_bao': [...]}"""
    t0 = time.time()
    os.environ['HC_FONTS'] = thu_muc_font
    try:
        import anh_mini; anh_mini.cai()                  # tren trinh duyet: giu anh de ve logo
    except Exception: pass
    wb = openpyxl.load_workbook(src, rich_text=True, keep_links=False)   # bo lien ket ngoai (doc rat cham, khong can de ve)
    ev = TinhCongThuc(wb); ev.tinh_het(); pdf = PDF(thu_muc_font); theme = _theme(wb); mdw = _mdw(wb, pdf)
    trang = []; canh = []
    for ws in wb.worksheets:
        if ws.sheet_state != 'visible': continue
        if log: log('@@PDF: %s' % ws.title)
        trang.append((ws.title, ve_sheet(ws, ev, pdf, theme, mdw, canh)))
    canh = ['Công thức không tính được (ô để trống trong PDF): ' + x for x in ev.loi] + list(dict.fromkeys(canh))
    n = pdf.luu(out)
    return {'trang': trang, 'canh_bao': canh, 'bytes': n, 'giay': round(time.time() - t0, 1), 'vi_tri': pdf.vt}

if __name__ == '__main__':
    fd = sys.argv[3] if len(sys.argv) > 3 else os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'fonts')
    R = xuat(sys.argv[1], sys.argv[2], fd)
    print(R['trang']); print('\n'.join(R['canh_bao'])); print('%d bytes, %.1fs' % (R['bytes'], R['giay']))
