# -*- coding: utf-8 -*-
"""Chuyen bao cao su co (BCSC) tu phuong/quan cu sang phuong moi, can cu TEN TU.
 - Doc tung bao cao (muc 3): "- Tu dieu khien: <ten tu>", "- Phuong <cu>", "- Quan <cu>".
 - Tra ten tu trong bang tra tu -> phuong moi (uu tien trung ten co dau, roi phuong thuoc quan cu).
 - Ghi PDF: SUA NOI TIEP (incremental update) - giu nguyen byte goc (chu ky so cua ban goc van kiem duoc),
   phu trang dong cu va viet: "- Phuong <moi>", "- Thanh pho Ho Chi Minh" (cung co chu, cung duong chan chu).
Dung: kq = chuyen(du_lieu_pdf, duong_dan_bang_tra, thu_muc_font)  -> {'pdf': bytes, 'bao_cao': [...], 'thong_ke': {...}}"""
import os, re, zlib, unicodedata
from bcsc import DocBCSC
from quan_phuong import phuong_cua_quan, quan_key

DONG_QUAN_MOI = 'Thành phố Hồ Chí Minh'
# Chuc danh o khoi ky (ben Cong ty CS) doi theo yeu cau anh Lam 09/11/2026
CHUC_DANH = [('PHỤ TRÁCH KỸ THUẬT THI CÔNG', 'CÁN BỘ PHỤ TRÁCH KỸ THUẬT')]
FONT_DAM = 'LiberationSerif-Bold.ttf'


def tim_chuc_danh(tl):
    """{so trang: [(cu, moi, x0, x1, y chan chu, co chu, dam?)]} - vi tri chuc danh can doi o khoi ky."""
    kq = {}
    for i, page in enumerate(tl.pdf.pages):
        for d in tl.dong(page):
            chu = ''.join(ch for ch, _ in d['chars'])
            for cu, moi in CHUC_DANH:
                j = chu.find(cu)
                if j < 0 or moi in chu: continue
                cs = [c for _, c in d['chars'][j:j + len(cu)]]
                kq.setdefault(i, []).append((cu, moi, cs[0]['x0'], cs[-1]['x1'], cs[0]['matrix'][5], cs[0]['size'],
                                             'Bold' in cs[0]['fontname']))
    return kq
FONT_CHU = 'LiberationSerif-Regular.ttf'          # cung do rong chu voi Times New Roman cua ban goc


# ------------------------------------------------------------------ tra ten tu
def _sa(s):
    s = unicodedata.normalize('NFD', str(s or '')).replace('đ', 'd').replace('Đ', 'D')
    return ''.join(c for c in s if unicodedata.category(c) != 'Mn')
def _norm(s): return re.sub(r'[^a-z0-9]', '', _sa(s).lower())                     # bo dau, bo khoang trang/gach
def _nguyen(s): return re.sub(r'[\s\-–]+', '', unicodedata.normalize('NFC', str(s or '')).lower())   # giu dau


def doc_bang_tra(path):
    os.environ.setdefault('MASTER_XLSX', path)             # refill_lib doc bien nay luc import
    import refill_lib as R
    R.MASTER = path
    return R.master_rows()                       # [(ten_tu, phuong, duong, dia_chi, goi)]


class TraTu:
    def __init__(self, rows):
        self.idx = {}
        for r in rows:
            if r[0] and r[1]: self.idx.setdefault(_norm(r[0]), []).append(r)

    def tra(self, ten_tu, quan_cu):
        """-> (phuong_moi | None, ghi_chu)"""
        if '\ufffd' in ten_tu: return None, 'không đọc được tên tủ (font lạ)'
        ds = self.idx.get(_norm(ten_tu), [])
        if not ds: return None, 'không có tủ này trong bảng tra'
        dung = [r for r in ds if _nguyen(r[0]) == _nguyen(ten_tu)]          # trung ca dau (Cừ ≠ Cự)
        if not dung:                                                         # chi trung khi bo dau: co the doc sai 1 dau -> khong tu doi
            return None, 'chỉ khớp khi bỏ dấu với tủ: ' + ', '.join(sorted({r[0] for r in ds}))
        ds = dung
        if len({r[1] for r in ds}) > 1:
            pq = set(phuong_cua_quan(quan_cu))
            q = [r for r in ds if r[1] in pq]
            if q: ds = q
        ph = sorted({r[1] for r in ds})
        if len(ph) > 1: return None, 'tủ trùng tên ở nhiều phường: ' + ', '.join(ph)
        return ph[0], ''


# ------------------------------------------------------------------ tim bao cao trong PDF
_TU = re.compile(r'^-\s*Tủ điều khiển\s*:\s*(.+)$')
_PH = re.compile(r'^-\s*(Phường|Xã|Thị trấn)\s+(.+)$')
_QU = re.compile(r'^-\s*(Quận|Huyện|Thành phố)\s+(.+)$')
_MA = re.compile(r'Mã sự cố\s*:\s*(\S+)')


def tim_bao_cao(tl):
    """[{trang, ma, tu, phuong(dong), quan(dong)}] - moi bao cao 1 muc. tl.ky_tu_la = so ky tu chua giai ma duoc."""
    kq = []; tl.ky_tu_la = 0; tl.anh_ky = {}
    for i, page, ds in tl.cac_trang():
        tl.ky_tu_la += sum(d['text'].count('\ufffd') for d in ds)
        # khoi ky: dong chuc danh "... PHU TRACH ..." -> anh chu ky tay nam giua dong chuc danh va dong ho ten
        cd = [d for d in ds if 'PHỤ TRÁCH' in d['text']]
        if cd and page.images:
            y0 = max(d['bottom'] for d in cd)
            sau = [d for d in ds if d['top'] > y0 + 15]
            y1 = min(d['top'] for d in sau) if sau else y0 + 160
            ten = {im['name'] for im in page.images if im['top'] >= y0 - 2 and im['bottom'] <= y1 + 5}
            if ten: tl.anh_ky[i] = ten
        ma = None
        for j, d in enumerate(ds):
            m = _MA.search(d['text'])
            if m: ma = m.group(1)
            m = _TU.match(d['text'])
            if not m: continue
            bc = dict(trang=i, page=page, ma=ma, tu=m.group(1).strip(), phuong=None, quan=None)
            for d2 in ds[j + 1:j + 6]:
                if bc['phuong'] is None and _PH.match(d2['text']): bc['phuong'] = d2
                elif bc['quan'] is None and _QU.match(d2['text']): bc['quan'] = d2
                if re.match(r'^\d+\s*-', d2['text']): break                     # sang muc 4
            kq.append(bc)
    return kq


# ------------------------------------------------------------------ ghi PDF noi tiep
def _ten(n): return '/' + ''.join(c if c.isalnum() or c in '-_.+' else '#%02X' % ord(c) for c in n)


def _pdf(o):
    """Doi tuong pdfminer -> byte cu phap PDF."""
    from pdfminer.psparser import PSLiteral, PSKeyword
    from pdfminer.pdftypes import PDFObjRef
    if o is None: return b'null'
    if o is True: return b'true'
    if o is False: return b'false'
    if isinstance(o, PDFObjRef): return b'%d 0 R' % o.objid
    if isinstance(o, PSLiteral): return _ten(o.name if isinstance(o.name, str) else o.name.decode('latin1')).encode('latin1')
    if isinstance(o, PSKeyword): return o.name if isinstance(o.name, bytes) else o.name.encode()
    if isinstance(o, int): return b'%d' % o
    if isinstance(o, float): return ('%.6f' % o).rstrip('0').rstrip('.').encode()
    if isinstance(o, (bytes, bytearray)): return b'<' + bytes(o).hex().encode() + b'>'
    if isinstance(o, str): return b'<' + o.encode('latin1', 'replace').hex().encode() + b'>'
    if isinstance(o, (list, tuple)): return b'[' + b' '.join(_pdf(x) for x in o) + b']'
    if isinstance(o, dict):
        return b'<<' + b''.join(_ten(k).encode('latin1') + b' ' + _pdf(v) for k, v in o.items()) + b'>>'
    raise TypeError('khong ghi duoc %r' % (o,))


class GhiNoiTiep:
    """Them doi tuong moi / ban moi cua doi tuong cu vao cuoi file (xref stream, /Prev)."""
    def __init__(self, goc, doc):
        self.goc = goc; self.doc = doc; self.obj = {}
        tr = doc.xrefs[0].trailer if doc.xrefs else {}
        self.size = max(int(x.trailer.get('Size', 0)) for x in doc.xrefs)
        self.trailer = {k: tr[k] for k in ('Root', 'Info', 'ID') if k in tr}
        for x in doc.xrefs:
            for k in ('Root', 'Info', 'ID'):
                if k not in self.trailer and k in x.trailer: self.trailer[k] = x.trailer[k]
        m = list(re.finditer(rb'startxref\s+(\d+)', goc[-2048:]))
        self.prev = int(m[-1].group(1))

    def moi(self, b):
        n = self.size; self.size += 1; self.obj[n] = b; return n

    def thay(self, n, b): self.obj[n] = b

    def stream(self, dic, data):
        data = zlib.compress(data, 6)
        return b'<<' + dic + b'/Filter/FlateDecode/Length %d>>stream\n' % len(data) + data + b'\nendstream'

    def luu(self):
        out = bytearray(self.goc)
        if not out.endswith(b'\n'): out += b'\n'
        vt = {}
        for n in sorted(self.obj):
            vt[n] = len(out); out += b'%d 0 obj\n' % n + self.obj[n] + b'\nendobj\n'
        xn = self.size; self.size += 1; vt[xn] = len(out)
        ns = sorted(vt); idx = []; i = 0
        while i < len(ns):
            j = i
            while j + 1 < len(ns) and ns[j + 1] == ns[j] + 1: j += 1
            idx += [ns[i], j - i + 1]; i = j + 1
        rows = b''.join(b'\x01' + vt[n].to_bytes(4, 'big') + b'\x00\x00' for n in ns)
        dic = (b'/Type/XRef/Size %d/W[1 4 2]/Index[%s]/Prev %d' % (self.size, ' '.join(map(str, idx)).encode(), self.prev)
               + b''.join(b'/' + k.encode() + b' ' + _pdf(v) for k, v in self.trailer.items()))
        out += b'%d 0 obj\n' % xn + self.stream(dic, rows) + b'\nendobj\nstartxref\n%d\n%%%%EOF\n' % vt[xn]
        return bytes(out)


def _font_obj(ttf, dung, ps='HCBCSC+LiberationSerif'):
    """(FontFile2, FontDescriptor, CIDFont, ToUnicode, Type0) dang byte; tham chieu giua chung la %(ff)d... de dien so doi tuong."""
    sub = ttf.subset(list(dung)); sc = 1000.0 / ttf.upm; gs = sorted(dung)
    fd = ('<</Type/FontDescriptor/FontName/%s/Flags 34/FontBBox[%d %d %d %d]/ItalicAngle 0/Ascent %d/Descent %d/CapHeight %d/StemV 80/FontFile2 %%(ff)d 0 R>>'
          % (ps, ttf.bbox[0] * sc, ttf.bbox[1] * sc, ttf.bbox[2] * sc, ttf.bbox[3] * sc, ttf.asc * sc, ttf.desc * sc, ttf.cap * sc))
    cid = ('<</Type/Font/Subtype/CIDFontType2/BaseFont/%s/CIDSystemInfo<</Registry(Adobe)/Ordering(Identity)/Supplement 0>>/FontDescriptor %%(fd)d 0 R/W[%s]/DW 1000/CIDToGIDMap/Identity>>'
           % (ps, ' '.join('%d[%d]' % (x, round(ttf.w[x] * sc)) for x in gs)))
    cm = ['/CIDInit/ProcSet findresource begin 12 dict begin begincmap/CIDSystemInfo<</Registry(Adobe)/Ordering(UCS)/Supplement 0>>def/CMapName/Adobe-Identity-UCS def/CMapType 2 def',
          '1 begincodespacerange <0000><FFFF> endcodespacerange', '%d beginbfchar' % len(gs)]
    cm += ['<%04X><%s>' % (x, dung[x].encode('utf-16-be').hex().upper()) for x in gs] + ['endbfchar', 'endcmap CMapName currentdict/CMap defineresource pop end end']
    t0 = '<</Type/Font/Subtype/Type0/BaseFont/%s/Encoding/Identity-H/DescendantFonts[%%(cid)d 0 R]/ToUnicode %%(tu)d 0 R>>' % ps
    return sub, fd, cid, '\n'.join(cm).encode(), t0


def _chuan_phuong(p):
    p = re.sub(r'\s+', ' ', unicodedata.normalize('NFC', str(p or ''))).strip()
    return p if not p or re.match(r'^(Phường|Xã|Thị trấn)\s', p) else 'Phường ' + p


def chuyen(du_lieu, bang_tra, thu_muc_font, log=print, kieu='sach', sua=None, nho=None, giu_ky=False):
    """bang_tra: duong dan (.xlsx/.json) hoac danh sach dong [(ten_tu, phuong, ...)].
    sua: {so trang (chuoi): {'tu': ten tu sua lai, 'phuong': phuong chon tay}} - nguoi dung sua tren app cho file nay.
    nho: [{'ten': ten tu, 'ten_goc': ten doc duoc tu PDF, 'phuong': ...}] - tu da nho tren may (uu tien hon bang tra).
    kieu='sach'    : file moi hoan toan - XOA chu cu, bo lich su sua va cac chu ky cu, giu o ky trong de ky lai.
    kieu='noi_tiep': ghi noi tiep, giu ban da ky ben trong (phu trang chu cu).
    giu_ky=True (voi kieu 'sach'): GIU HINH KY - giu hinh chu ky tay, ve hinh chu ky so thang len trang;
        van go o/du lieu chu ky so (chu ky so khong con hieu luc sau khi sua)."""
    from pdfminer.pdftypes import resolve1
    import xuat_pdf
    rows = doc_bang_tra(bang_tra) if isinstance(bang_tra, str) else bang_tra
    tra = TraTu(rows)
    tl = DocBCSC(du_lieu)
    bcs = tim_bao_cao(tl)
    ttf = xuat_pdf.TTF(open(os.path.join(thu_muc_font, FONT_CHU), 'rb').read(), FONT_CHU)
    dung = {}                                    # gid -> ky tu
    def hexs(s):
        o = []
        for ch in unicodedata.normalize('NFC', s):
            g = ttf.cmap.get(ord(ch))
            if not g: raise ValueError('font không có chữ "%s" (trong "%s")' % (ch, s))
            dung[g] = ch; o.append('%04X' % g)
        return '<' + ''.join(o) + '>'

    bao = []; sua_trang = {}                    # so trang -> [page, [lenh ve], [(x, y chan chu) dong cu can xoa]]
    sua = sua or {}
    nho_goc, nho_ten = {}, {}
    for x in (nho or []):
        if x.get('phuong'):
            if x.get('ten_goc'): nho_goc[x['ten_goc']] = _chuan_phuong(x['phuong'])
            if x.get('ten'): nho_ten[_nguyen(x['ten'])] = _chuan_phuong(x['phuong'])
    for bc in bcs:
        ph_cu = _PH.match(bc['phuong']['text']).group(2).strip() if bc['phuong'] else ''
        q_cu = _QU.match(bc['quan']['text']).group(2).strip() if bc['quan'] else ''
        k = str(bc['trang'] + 1); tu = bc['tu']; nguon = ''
        if k in sua and sua[k].get('phuong'):                       # sua tay tren app
            tu = (sua[k].get('tu') or tu).strip(); ph_moi, ghi, nguon = _chuan_phuong(sua[k]['phuong']), '', 'sửa tay'
        elif bc['tu'] in nho_goc or _nguyen(bc['tu']) in nho_ten:      # tu da nho tren may
            ph_moi, ghi, nguon = nho_goc.get(bc['tu']) or nho_ten[_nguyen(bc['tu'])], '', 'đã nhớ'
        else:
            ph_moi, ghi = tra.tra(bc['tu'], q_cu)
        r = dict(trang=bc['trang'] + 1, ma=bc['ma'], tu=tu, tu_goc=bc['tu'], phuong_cu=ph_cu, quan_cu=q_cu, phuong_moi=ph_moi or '', ghi_chu=ghi, nguon=nguon)
        if not bc['phuong']: r['ghi_chu'] = 'không thấy dòng Phường'; ph_moi = None
        bao.append(r)
        if not ph_moi: continue
        ph_moi_ten = re.sub(r'^(Phường|Xã|Thị trấn)\s+', '', ph_moi)
        loai = re.match(r'^(Phường|Xã|Thị trấn)', ph_moi)
        moi = [(bc['phuong'], '- %s %s' % (loai.group(1) if loai else 'Phường', ph_moi_ten))]
        if bc['quan']: moi.append((bc['quan'], '- ' + DONG_QUAN_MOI))
        page = bc['page']; H = float(page.mediabox[3]); y0 = float(page.mediabox[1])
        e = sua_trang.setdefault(bc['trang'], [page, [], []]); L = e[1]
        for d, s in moi:
            c0 = d['chars'][0][1]; x, yb = c0['matrix'][4], c0['matrix'][5]; sz = c0['size']
            e[2].append((x, yb))
            if kieu == 'noi_tiep':
                top = H + y0 - min(c['top'] for _, c in d['chars']); bot = H + y0 - max(c['bottom'] for _, c in d['chars'])
                x1 = max(d['x1'], x + ttf.rong(s, sz))
                L.append('1 1 1 rg %.2f %.2f %.2f %.2f re f' % (x - 1, bot - 0.5, x1 - x + 2.5, top - bot + 1))
            L.append('BT /HCBC %.2f Tf 0 0 0 rg 1 0 0 1 %.2f %.2f Tm %s Tj ET' % (sz, x, yb, hexs(s)))

    ttfb = xuat_pdf.TTF(open(os.path.join(thu_muc_font, FONT_DAM), 'rb').read(), FONT_DAM)
    dungb = {}
    def hexb(s):
        o = []
        for ch in unicodedata.normalize('NFC', s):
            g = ttfb.cmap.get(ord(ch))
            if not g: raise ValueError('font không có chữ "%s"' % ch)
            dungb[g] = ch; o.append('%04X' % g)
        return '<' + ''.join(o) + '>'
    n_cd = 0
    if kieu == 'sach':
        for so, ds in tim_chuc_danh(tl).items():
            e = sua_trang.setdefault(so, [tl.pdf.pages[so], [], []])
            for cu, moi, x0, x1, yb, sz, dam in ds:
                f, fn = (ttfb, 'HCBD') if dam else (ttf, 'HCBC')
                xm = (x0 + x1) / 2 - f.rong(moi, sz) / 2                       # canh giua dung cho chuc danh cu
                e[2].append((x0, yb)); n_cd += 1
                e[1].append('BT /%s %.2f Tf 0 0 0 rg 1 0 0 1 %.2f %.2f Tm %s Tj ET' % (fn, sz, xm, yb, hexb(moi) if dam else hexs(moi)))
    tk = dict(ky_tu_la=getattr(tl, 'ky_tu_la', 0), chuc_danh=n_cd, bao_cao=len(bao), da_doi=sum(1 for r in bao if r['phuong_moi'] and not r['ghi_chu']),
              khong_doi=sum(1 for r in bao if not r['phuong_moi'] or r['ghi_chu']))
    if not sua_trang:
        return dict(pdf=du_lieu, bao_cao=bao, thong_ke=tk, chu_ky_bo=0, anh_ky_bo=0, ky_giu=0, che_do='giu_ky' if giu_ky else 'bo_ky')
    if kieu == 'sach':
        fonts = [('/HCBC', _font_obj(ttf, dung))] if dung else []
        if dungb: fonts.append(('/HCBD', _font_obj(ttfb, dungb, 'HCBCSD+LiberationSerif-Bold')))
        if giu_ky:
            out, bo, phang = _ghi_sach(du_lieu, sua_trang, fonts, None, giu_hinh=True)
            giu = sum(len(v) for v in getattr(tl, 'anh_ky', {}).values()) + phang
            log('BCSC: %d báo cáo, đổi %d, không đổi %d; giữ %d hình ký (gỡ %d chữ ký số cũ)' % (tk['bao_cao'], tk['da_doi'], tk['khong_doi'], giu, bo))
            return dict(pdf=out, bao_cao=bao, thong_ke=tk, chu_ky_bo=bo, anh_ky_bo=0, ky_giu=giu, che_do='giu_ky')
        out, bo, anh = _ghi_sach(du_lieu, sua_trang, fonts, getattr(tl, 'anh_ky', {}))
        log('BCSC: %d báo cáo, đổi %d, không đổi %d; bỏ %d chữ ký số cũ, %d hình chữ ký' % (tk['bao_cao'], tk['da_doi'], tk['khong_doi'], bo, anh))
        return dict(pdf=out, bao_cao=bao, thong_ke=tk, chu_ky_bo=bo, anh_ky_bo=anh, ky_giu=0, che_do='bo_ky')

    sub, fd_s, cid_s, tu_s, t0_s = _font_obj(ttf, dung)
    g = GhiNoiTiep(du_lieu, tl.pdf.doc)
    ff = g.moi(g.stream(b'/Length1 %d' % len(sub), sub))
    fd = g.moi((fd_s % dict(ff=ff)).encode())
    cid = g.moi((cid_s % dict(fd=fd)).encode())
    tu = g.moi(g.stream(b'', tu_s))
    fo = g.moi((t0_s % dict(cid=cid, tu=tu)).encode())
    q_mo = g.moi(g.stream(b'', b'q'))
    from pdfminer.pdftypes import PDFObjRef
    for so in sorted(sua_trang):
        page, L, _ = sua_trang[so]
        po = page.page_obj
        dic = dict(po.attrs)
        res = dict(resolve1(po.resources) or {})                                 # tai nguyen (ke ca ke thua)
        fr = dict(resolve1(res.get('Font', {})) or {})
        fr['HCBC'] = PDFObjRef(None, fo, 0)
        res['Font'] = fr; dic['Resources'] = res
        for k in ('MediaBox', 'Rotate'):                                        # thuoc tinh ke thua: ghi ro
            if k not in dic and getattr(po, k.lower(), None) is not None: dic[k] = getattr(po, k.lower())
        nd = g.moi(g.stream(b'', ('Q q\n' + '\n'.join(L) + '\nQ\n').encode()))
        c = dic.get('Contents'); c = resolve1(c) if not isinstance(c, PDFObjRef) else c
        goc = list(c) if isinstance(c, list) else ([c] if c is not None else [])
        dic['Contents'] = [PDFObjRef(None, q_mo, 0)] + goc + [PDFObjRef(None, nd, 0)]
        g.thay(po.pageid, _pdf(dic))
    out = g.luu()
    log('BCSC: %d báo cáo, đổi %d, không đổi %d' % (tk['bao_cao'], tk['da_doi'], tk['khong_doi']))
    return dict(pdf=out, bao_cao=bao, thong_ke=tk, chu_ky_bo=0)


def _mul(a, b):
    return (a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3],
            a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5])


def _loc_chu(ops, xoa):
    """Bo cac lenh hien chu (Tj TJ ' ") co diem goc trung (x, y chan chu) cua dong cu. Tra (ops moi, [so lenh bo moi dong])."""
    I = (1, 0, 0, 1, 0, 0); ctm = I; st = []; tm = tlm = I; TL = 0.0; kq = []; bo = [0] * len(xoa)
    for opd, op in ops:
        f = [float(v) for v in opd] if op in (b'cm', b'Tm', b'Td', b'TD', b'TL') else None
        if op == b'q': st.append(ctm)
        elif op == b'Q': ctm = st.pop() if st else I
        elif op == b'cm': ctm = _mul(tuple(f), ctm)
        elif op == b'BT': tm = tlm = I
        elif op == b'Tm': tm = tlm = tuple(f)
        elif op in (b'Td', b'TD'):
            if op == b'TD': TL = -f[1]
            tlm = _mul((1, 0, 0, 1, f[0], f[1]), tlm); tm = tlm
        elif op == b'TL': TL = f[0]
        elif op in (b'T*', b"'", b'"'):
            tlm = _mul((1, 0, 0, 1, 0, -TL), tlm); tm = tlm
        if op in (b'Tj', b'TJ', b"'", b'"'):
            m = _mul(tm, ctm)
            hit = [i for i, (x, y) in enumerate(xoa) if abs(m[5] - y) < 0.6 and m[4] > x - 1.5]
            if hit:
                bo[hit[0]] += 1
                if op == b'"': kq += [([opd[0]], b'Tw'), ([opd[1]], b'Tc')]   # giu tac dung phu cua "
                if op in (b"'", b'"'): kq.append(([], b'T*'))                  # giu viec xuong dong
                continue
        kq.append((opd, op))
    return kq, bo


def _ghi_sach(du_lieu, sua_trang, fonts, anh_ky=None, giu_hinh=False):
    """Ghi lai TOAN BO file (pypdf): xoa chu cu, them chu moi, bo toan bo o/chu ky so, khong con lich su sua.
    giu_hinh=True: truoc khi go o chu ky so da ky, ve hinh hien thi (/AP /N) cua no thang len trang.
    Tra (bytes, so chu ky so da go, so hinh: anh tay da bo | hinh chu ky so da ve len trang)."""
    import io
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import (ContentStream, DecodedStreamObject, DictionaryObject, NameObject, ArrayObject,
                               NumberObject, IndirectObject)
    rd = PdfReader(io.BytesIO(du_lieu), strict=False)
    w = PdfWriter(clone_from=rd)
    for pg in w.pages:                                          # BCSC: ToUnicode tro toi doi tuong khong ton tai -> bo
        try:
            fr = pg['/Resources'].get_object().get('/Font')
            for k, v in (fr.get_object().items() if fr is not None else []):
                f = v.get_object()
                if '/ToUnicode' in f:
                    try: ok = f['/ToUnicode'].get_object() is not None and hasattr(f['/ToUnicode'].get_object(), 'get_data')
                    except Exception: ok = False
                    if not ok: del f['/ToUnicode']
        except Exception:
            pass

    def st(dic, data):
        s = DecodedStreamObject(); s.set_data(data)
        for k, v in dic.items(): s[NameObject(k)] = v
        return w._add_object(s)
    def tu_chuoi(b, **ref):                                    # dict PDF dang chuoi -> doi tuong pypdf
        from pypdf.generic import read_object
        txt = re.sub(r'%\((\w+)\)d 0 R', 'null', b)            # tham chieu dien sau ('0 0 R' lam pypdf doc lech, mat /W)
        o = read_object(io.BytesIO(txt.encode()), None)
        return o
    fo = {}
    for ten, (sub, fd_s, cid_s, tu_s, t0_s) in fonts:
        ff = st({'/Length1': NumberObject(len(sub))}, sub)
        fd = tu_chuoi(fd_s, ff=1); fd[NameObject('/FontFile2')] = ff; fd = w._add_object(fd)
        cid = tu_chuoi(cid_s, fd=1); cid[NameObject('/FontDescriptor')] = fd; cid = w._add_object(cid)
        tu = st({}, tu_s)
        t0 = tu_chuoi(t0_s, cid=1, tu=1); t0[NameObject('/DescendantFonts')] = ArrayObject([cid]); t0[NameObject('/ToUnicode')] = tu
        fo[ten] = w._add_object(t0)

    for so, (page, L, xoa) in sorted(sua_trang.items()):
        pg = w.pages[so]
        cs = ContentStream(pg.get_contents(), w)
        ops, n = _loc_chu(cs.operations, xoa)
        if min(n) == 0:
            raise ValueError('trang %d: không tìm thấy dòng chữ cũ để xoá (%s)' % (so + 1, n))
        cs.operations = ops
        data = b'q\n' + cs.get_data() + b'\nQ\n' + '\n'.join(L).encode() + b'\n'
        pg[NameObject('/Contents')] = st({}, data)
        res = pg.get('/Resources')
        res = res.get_object() if res is not None else DictionaryObject()
        fr = res.get('/Font'); fr = fr.get_object() if fr is not None else DictionaryObject()
        fr = DictionaryObject(fr)
        for ten, ref in fo.items(): fr[NameObject(ten)] = ref
        res = DictionaryObject(res); res[NameObject('/Font')] = fr
        pg[NameObject('/Resources')] = res

    # bo TOAN BO o chu ky so (da ky va o trong); bo dem so chu ky da ky
    bo = 0; ref_bo = set(); xoa_obj = set(); anh = 0
    # bo hinh chu ky tay o khoi ky (anh nam giua dong chuc danh va dong ho ten)
    for so, ten in sorted((anh_ky or {}).items()):
        pg = w.pages[so]; ten = {'/' + t for t in ten}
        cs = ContentStream(pg.get_contents(), w)
        ops = [(o, op) for o, op in cs.operations if not (op == b'Do' and o and str(o[0]) in ten)]
        if len(ops) == len(cs.operations): continue
        anh += len(cs.operations) - len(ops); cs.operations = ops
        pg[NameObject('/Contents')] = st({}, cs.get_data())
        res = DictionaryObject(pg['/Resources'].get_object())
        xo = DictionaryObject(res['/XObject'].get_object())
        for t in ten:
            if t in xo:
                r_ = xo.raw_get(t)
                if isinstance(r_, IndirectObject): xoa_obj.add(r_.idnum)
                del xo[t]
        res[NameObject('/XObject')] = xo; pg[NameObject('/Resources')] = res
    phang = 0
    if giu_hinh:                                                # ve hinh chu ky so (da ky) thang len trang
        for so, pg in enumerate(w.pages):
            an = pg.get('/Annots')
            if an is None: continue
            ve = []
            for a in an.get_object():
                o = a.get_object(); par = o.get('/Parent'); fld = par.get_object() if par is not None else o
                if fld.get('/FT') != '/Sig' or fld.get('/V') is None: continue          # o ky trong: khong ve
                if int(o.get('/F', 0)) & 2: continue                                    # an
                ap = o.get('/AP'); n = ap.get_object().get('/N') if ap is not None else None
                if n is None or '/Rect' not in o: continue
                ns = n.get_object()
                if not hasattr(ns, 'get_data'): continue
                bb = [float(v) for v in ns.get('/BBox', [0, 0, 0, 0])]
                mx = [float(v) for v in ns.get('/Matrix', [1, 0, 0, 1, 0, 0])]
                pts = [(x * mx[0] + y * mx[2] + mx[4], x * mx[1] + y * mx[3] + mx[5]) for x in (bb[0], bb[2]) for y in (bb[1], bb[3])]
                bx0, by0 = min(p[0] for p in pts), min(p[1] for p in pts); bx1, by1 = max(p[0] for p in pts), max(p[1] for p in pts)
                rc = [float(v) for v in o['/Rect']]; rx0, ry0, rx1, ry1 = min(rc[0], rc[2]), min(rc[1], rc[3]), max(rc[0], rc[2]), max(rc[1], rc[3])
                if bx1 - bx0 <= 0 or by1 - by0 <= 0: continue
                sx, sy = (rx1 - rx0) / (bx1 - bx0), (ry1 - ry0) / (by1 - by0)
                if '/Type' not in ns: ns[NameObject('/Type')] = NameObject('/XObject')
                if '/Subtype' not in ns: ns[NameObject('/Subtype')] = NameObject('/Form')
                ref = n if isinstance(n, IndirectObject) else w._add_object(ns)
                ve.append(('/HCKy%d' % len(ve), ref, (sx, 0, 0, sy, rx0 - bx0 * sx, ry0 - by0 * sy)))
            if not ve: continue
            res = DictionaryObject(pg['/Resources'].get_object()) if '/Resources' in pg else DictionaryObject()
            xo = res.get('/XObject'); xo = DictionaryObject(xo.get_object()) if xo is not None else DictionaryObject()
            lenh = []
            for ten, ref, m in ve:
                xo[NameObject(ten)] = ref
                lenh.append('q %.4f %.4f %.4f %.4f %.4f %.4f cm %s Do Q' % (m + (ten,)))
            res[NameObject('/XObject')] = xo; pg[NameObject('/Resources')] = res
            cs = ContentStream(pg.get_contents(), w)
            pg[NameObject('/Contents')] = st({}, b'q\n' + cs.get_data() + b'\nQ\n' + '\n'.join(lenh).encode() + b'\n')
            phang += len(ve)
    af = w._root_object.get('/AcroForm')
    if af is not None:
        af = af.get_object(); giu = ArrayObject()
        for fr_ in af.get('/Fields', []):
            f = fr_.get_object()
            if f.get('/FT') == '/Sig':
                if f.get('/V') is not None: bo += 1
                v = f.raw_get('/V') if '/V' in f else None
                if isinstance(v, IndirectObject): xoa_obj.add(v.idnum)
                if isinstance(fr_, IndirectObject): ref_bo.add(fr_.idnum); xoa_obj.add(fr_.idnum)
                for k in f.get('/Kids', []): ref_bo.add(k.idnum); xoa_obj.add(k.idnum)
                ap = f.get('/AP')                                               # hinh chu ky (appearance)
                if ap is not None and not giu_hinh:
                    for vv in ap.get_object().values():
                        if isinstance(vv, IndirectObject): xoa_obj.add(vv.idnum)
            else: giu.append(fr_)
        if len(giu): af[NameObject('/Fields')] = giu
        else: del w._root_object['/AcroForm']
        if '/SigFlags' in af: del af['/SigFlags']
        def la_ky(a):
            if isinstance(a, IndirectObject) and a.idnum in ref_bo: return True
            o = a.get_object(); par = o.get('/Parent')
            return o.get('/FT') == '/Sig' or (par is not None and par.get_object().get('/FT') == '/Sig')
        for pg in w.pages:
            an = pg.get('/Annots')
            if an is None: continue
            an = an.get_object()
            moi = ArrayObject([a for a in an if not la_ky(a)])
            if len(moi) != len(an):
                if len(moi): pg[NameObject('/Annots')] = moi
                else: del pg['/Annots']
    if '/Perms' in w._root_object: del w._root_object['/Perms']        # DocMDP tro toi chu ky da bo
    from pypdf.generic import NullObject
    for n in xoa_obj:                                                          # khong de sot chu ky cu trong file
        if 0 < n <= len(w._objects): w._objects[n - 1] = NullObject()
    for pg in w.pages:
        try: pg.compress_content_streams()
        except Exception: pass
    b = io.BytesIO(); w.write(b)
    return b.getvalue(), bo, (phang if giu_hinh else anh)


def xuat_nho(nho, out):
    """Xuat danh sach tu da nho tren may ra Excel (de bo sung bang tra chinh thuc)."""
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Tu bo sung'
    tieu = ['STT', 'Tên tủ điều khiển', 'Phường/Xã mới', 'Tên đọc từ PDF', 'Ngày nhớ', 'File nguồn']
    ws.append(tieu)
    for i, x in enumerate(nho, 1):
        ws.append([i, x.get('ten', ''), _chuan_phuong(x.get('phuong', '')), x.get('ten_goc', '') if x.get('ten_goc') != x.get('ten') else '',
                   x.get('ngay', ''), x.get('file', '')])
    m = Side(style='thin'); b = Border(left=m, right=m, top=m, bottom=m)
    for row in ws.iter_rows():
        for c in row:
            c.border = b; c.font = Font(name='Times New Roman', size=12, bold=c.row == 1)
            c.alignment = Alignment(vertical='center', wrap_text=True, horizontal='center' if c.column in (1, 5) or c.row == 1 else 'left')
    for col, w in zip('ABCDEF', (6, 34, 24, 30, 13, 36)): ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'
    wb.save(out)
