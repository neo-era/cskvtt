# -*- coding: utf-8 -*-
"""Nhat ky thi cong hang ngay (NKTC, Crystal Reports): dien cac o de trong cua tung ngay.
 - Nhiet do sang/chieu: tu thoi tiet that (Open-Meteo, trang web lay theo gio) -> thoi_tiet().
 - Cong nhan, thiet bi: tinh theo dinh muc (file "Ma cong viec" .xlsb/.xlsx, bang Phan tich don gia)
   = tong (khoi luong x dinh muc nhan cong/may) cac viec trong ngay, toi thieu 1 to: 04 nguoi + 01 xe thang.
 - Tick "Cong tac ve sinh moi truong" va "An toan lao dong, an toan giao thong" (mac dinh Binh thuong).
Nguoi dung xem/sua tung ngay tren app roi moi xuat (de_xuat -> sua -> dien)."""
import io, math, os, re, unicodedata

TO_NGUOI, TO_XE = 4, 1                                  # toi thieu 1 to
TRAN_NGUOI, TRAN_XE = 12, 3                             # nang luc toi da/ngay cua don vi (anh Lam chot 13/11/2026: 3 to)
MAY_NHO = 0.1                                           # may phu < 0,1 ca/ngay: khong ghi vao nhat ky
TICK = ('tot', 'binh_thuong', 'kem')
_NGAY = re.compile(r'^Ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})\s*$')
_VIEC = re.compile(r'^•\s*(.+):\s*([\d.,]+)\s+(.+?)\s*$')     # don vi co the nhieu chu: "đầu cáp", "1 Trụ"


def _k(s):
    s = unicodedata.normalize('NFD', str(s or '').lower()).replace('đ', 'd')
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return re.sub(r'[^a-z0-9]+', ' ', s).strip()


def so_kl(s):
    """Khoi luong kieu VN: dau phay thap phan, dau cham ngan cach hang nghin."""
    return float(str(s).strip().replace('.', '').replace(',', '.'))


# ---------------- Dinh muc ----------------
def _cac_dong(du_lieu, ten_file):
    """Tra (ten sheet, dong) cua bang Phan tich don gia."""
    ten = (ten_file or '').lower()
    try:
        if ten.endswith('.xlsb'):
            from pyxlsb import open_workbook
            wb = open_workbook(io.BytesIO(du_lieu))
            ds = [s for s in wb.sheets if _k(s).startswith('phan tich dg')]
            if not ds: raise ValueError('không có sheet "Phan tich DG…"')
            sh = sorted(ds, key=lambda s: (0 if 'kv1' in _k(s) else 1, s))[0]
            with wb.get_sheet(sh) as ws:
                return sh, [[c.v for c in r] for r in ws.rows()]
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(du_lieu), read_only=True, data_only=True)
        ds = [s for s in wb.sheetnames if _k(s).startswith('phan tich dg')]
        if not ds: raise ValueError('không có sheet "Phan tich DG…"')
        sh = sorted(ds, key=lambda s: (0 if 'kv1' in _k(s) else 1, s))[0]
        return sh, [list(r) for r in wb[sh].iter_rows(values_only=True)]
    except ValueError as e:
        raise ValueError('File định mức không đúng: %s' % e)
    except Exception as e:
        raise ValueError('Không đọc được file định mức (Mã công việc): %s' % e)


def doc_dinh_muc(du_lieu, ten_file):
    """-> {'nguon': sheet, 'muc': {khoa ten: {'ten','dv','nc' (cong/dv), 'may' {ten may: ca/dv}}}}"""
    sh, dong = _cac_dong(du_lieu, ten_file)
    muc = {}; cur = None
    for r in dong:
        v = [x for x in r if x not in (None, '')]
        if not v: continue
        hm = [i for i, x in enumerate(v) if isinstance(x, str) and re.fullmatch(r'HM\d+(_\w+)?', x.strip())]
        if hm and hm[0] >= 1 and hm[0] + 2 < len(v) and isinstance(v[hm[0] + 1], str) and isinstance(v[hm[0] + 2], str):
            i = hm[0]
            cur = {'ten': v[i + 1].strip(), 'dv': v[i + 2].strip(), 'nc': 0.0, 'may': {}}
            k = _k(cur['ten'])
            if k and k not in muc: muc[k] = cur
            else: cur = None if k in muc and muc[k] is not cur else cur
            continue
        if cur is None or len(v) < 5 or not isinstance(v[3], str): continue
        dv = v[3].strip()
        if dv not in ('công', 'ca'): continue
        try: d = float(v[4])
        except (TypeError, ValueError): continue
        if dv == 'công': cur['nc'] = round(cur['nc'] + d, 6)
        else:
            t = re.sub(r'\s+', ' ', str(v[2]).replace('\xa0', ' ')).strip()
            cur['may'][t] = round(cur['may'].get(t, 0) + d, 6)
    if len(muc) < 10:
        raise ValueError('File định mức không đúng: sheet "%s" không có bảng phân tích đơn giá' % sh)
    return {'nguon': sh, 'muc': muc}


def tra_dinh_muc(dm, ten):
    m = dm['muc'].get(_k(ten))
    if m is None:                                         # bo ghi chu trong ngoac cuoi ten
        m = dm['muc'].get(_k(re.sub(r'\s*\([^()]*\)\s*$', '', ten)))
    return m


def _ten_may(t):
    t = re.sub(r'ôtô', 'ô tô', t, flags=re.I)
    t = re.split(r'\s+-\s+|\s+sức nâng|,', t)[0].strip()
    return t[:1].lower() + t[1:]


def tinh_nhan_luc(viec, dm, tran=None):
    """viec: [(ten, khoi luong, don vi)] -> cong_nhan, thiet_bi (chuoi), cong (tong cong), thieu (viec khong co dinh muc),
    cong_nhan_dm / xe_dm (theo dung dinh muc, chua chan), vuot (dinh muc vuot nang luc -> da chan o tran),
    chi_cung_cap (ngay chi co viec cung cap vat tu/dich vu: khong ghi nhan luc, thiet bi)."""
    tn, tx = tran or (TRAN_NGUOI, TRAN_XE)
    cong = 0.0; nang = 0.0; khac = {}; thieu = []
    for ten, kl, dv in viec:
        m = tra_dinh_muc(dm, ten)
        if m is None:
            thieu.append(ten); continue
        cong += kl * m['nc']
        for mt, ca in m['may'].items():
            if re.search(r'xe (thang|nang)', _k(mt)): nang += kl * ca
            else:
                t = _ten_may(mt); khac[t] = khac.get(t, 0) + kl * ca
    if viec and not thieu and cong <= 1e-9 and nang <= 1e-9 and not any(v > 1e-9 for v in khac.values()):
        return dict(cong_nhan=0, cong=0.0, thiet_bi='', thieu=[], cong_nhan_dm=0, xe_dm=0, vuot=False, chi_cung_cap=True)
    cn_dm = max(TO_NGUOI, math.ceil(cong - 1e-9)); xe_dm = max(TO_XE, math.ceil(nang - 1e-9))
    cn, xe = (min(cn_dm, tn) if tn else cn_dm), (min(xe_dm, tx) if tx else xe_dm)
    tb = ['%02d xe thang' % xe]
    tb += ['%02d %s' % (max(1, math.ceil(v - 1e-9)), t) for t, v in sorted(khac.items()) if v >= MAY_NHO - 1e-9]
    return dict(cong_nhan=cn, cong=round(cong, 6), thiet_bi=', '.join(tb), thieu=thieu, cong_nhan_dm=cn_dm, xe_dm=xe_dm,
                ca_xe=round(nang, 4), vuot=cn < cn_dm or xe < xe_dm, chi_cung_cap=False)


# ---------------- Thoi tiet (Open-Meteo, theo gio) ----------------
def thoi_tiet(js, sang=(7, 11), chieu=(13, 17), nguong_mua=0.5):
    """-> {'YYYY-MM-DD': {'sang','chieu' (vd '29°C – 31°C'), 'mua_sang','mua_chieu'}}"""
    h = (js or {}).get('hourly') or {}
    tg, nd, mu = h.get('time') or [], h.get('temperature_2m') or [], h.get('precipitation') or []
    ngay = {}
    for i, t in enumerate(tg):
        try: d, gio = t[:10], int(t[11:13])
        except (ValueError, TypeError): continue
        for buoi, (a, b) in (('sang', sang), ('chieu', chieu)):
            if a <= gio <= b:
                x = ngay.setdefault(d, {}).setdefault(buoi, ([], []))
                if i < len(nd) and nd[i] is not None: x[0].append(float(nd[i]))
                if i < len(mu) and mu[i] is not None: x[1].append(float(mu[i]))
    tron = lambda v: int(math.floor(v + 0.5))
    kq = {}
    for d, bs in ngay.items():
        o = {}
        for buoi in ('sang', 'chieu'):
            t, m = bs.get(buoi, ([], []))
            if t:
                lo, hi = tron(min(t)), tron(max(t))
                o[buoi] = '%d°C' % lo if lo == hi else '%d°C – %d°C' % (lo, hi)
            else: o[buoi] = ''
            o['mua_' + buoi] = sum(m) >= nguong_mua
        if o['sang'] or o['chieu']: kq[d] = o
    return kq


# ---------------- Doc file nhat ky ----------------
def _y(c):
    return c['matrix'][5]


def doc_nhat_ky(du_lieu):
    """-> [ {ngay 'YYYY-MM-DD', trang [chi so], muc [{ten, kl, dv}], vi_tri {...}} ]"""
    import pdfplumber
    try:
        pdf = pdfplumber.open(io.BytesIO(du_lieu)); trang = list(pdf.pages)
    except Exception as e:
        raise ValueError('Không đọc được file PDF: %s' % e)
    if not trang: raise ValueError('File PDF rỗng')
    ngay = []; cur = None
    for so, pg in enumerate(trang):
        try: dong = pg.extract_text_lines()
        except Exception as e: raise ValueError('Không đọc được file PDF (trang %d): %s' % (so + 1, e))
        tde = [i for i, l in enumerate(dong) if l['text'].strip() == 'NHẬT KÝ THI CÔNG HÀNG NGÀY' and i + 1 < len(dong) and _NGAY.match(dong[i + 1]['text'].strip())]
        if len(tde) > 1 or (tde and tde[0] > 2):
            raise ValueError('Trang %d có ngày mới bắt đầu giữa trang — mẫu nhật ký này chưa hỗ trợ, gửi file để bổ sung' % (so + 1))
        for i in tde:
            g = _NGAY.match(dong[i + 1]['text'].strip())
            cur = {'ngay': '%s-%02d-%02d' % (g.group(3), int(g.group(2)), int(g.group(1))), 'trang': [], 'muc': [], 'vi_tri': {}, '_ten': set(), '_tick': [], 'doc_loi': []}
            ngay.append(cur)
        if cur is None:
            continue
        cur['trang'].append(so)
        vt = cur['vi_tri']
        for i, l in enumerate(dong):
            t = l['text'].strip(); cs = l['chars']
            if t.startswith('•'):
                m = _VIEC.match(t)
                if not m and i + 1 < len(dong) and not dong[i + 1]['text'].strip().startswith(('•', '-')):   # ten viec xuong dong
                    m = _VIEC.match(t + ' ' + dong[i + 1]['text'].strip())
                if not m: cur['doc_loi'].append(t)
                elif m.group(1) not in cur['_ten']:
                    cur['_ten'].add(m.group(1))
                    try: cur['muc'].append({'ten': m.group(1).strip(), 'kl': so_kl(m.group(2)), 'dv': m.group(3)})
                    except ValueError: cur['doc_loi'].append(t)
            if 'Sáng' in t and 'Chiều' in t and 'cot' not in vt:                 # tieu de cot bang thoi tiet
                a, b = _cum(l, 'Sáng'), _cum(l, 'Chiều')
                if a and b: vt['cot'] = [(a[0][0] + a[0][1]) / 2, (b[0][0] + b[0][1]) / 2]
            if t.startswith('Thời tiết') and 'thoi_tiet' not in vt:
                nh = _cum(l, 'Thời tiết'); sau = [c for c in cs if c['text'].strip() and nh and c['x0'] > nh[0][1] + 3]
                vt['thoi_tiet'] = {'trang': so, 'y': _y(cs[0]), 'size': cs[0]['size'], 'x_xoa': sau[0]['x0'] if sau else None}
            elif t.startswith('Nhiệt độ') and 'nhiet_do' not in vt:
                vt['nhiet_do'] = {'trang': so, 'y': _y(cs[0]), 'size': cs[0]['size']}
            elif t.startswith('- Công nhân:') and 'cong_nhan' not in vt:
                w = [c for c in cs if c['text'] == ':']
                ng = _cum(l, 'người')
                if w and ng: vt['cong_nhan'] = {'trang': so, 'y': _y(cs[0]), 'size': cs[0]['size'], 'x': (w[0]['x1'] + ng[0][0]) / 2}
            elif t.startswith('2.2. Thiết bị thi công') and 'thiet_bi' not in vt:
                vt['thiet_bi'] = {'trang': so, 'y': _y(cs[0]), 'size': cs[0]['size'], 'x': l['x1'] + 6, 'x_max': float(pg.width) - 24}
            elif re.fullmatch(r'Tốt\s+Bình thường\s+Kém', t):     # hang o tick (muc 4 roi muc 5; co the sang trang)
                o = _o_tick(pg, l, so)
                if len(o) == 3: cur['_tick'].append(o)
    if not ngay:
        raise ValueError('Không phải file nhật ký thi công (không thấy tiêu đề "NHẬT KÝ THI CÔNG HÀNG NGÀY")')
    if ngay[0]['trang'][0] != 0:                          # trang le truoc ngay dau: gan vao ngay dau
        ngay[0]['trang'] = list(range(ngay[0]['trang'][0])) + ngay[0]['trang']
    for n in ngay:
        n.pop('_ten', None); tk = n.pop('_tick'); vt = n['vi_tri']
        for k, o in zip(('vsmt', 'atld'), tk): vt[k] = o
        n['thieu_o'] = [ten for k, ten in (('cot', 'cột Sáng/Chiều'), ('nhiet_do', 'dòng Nhiệt độ'), ('cong_nhan', 'dòng Công nhân'),
                                            ('thiet_bi', 'dòng Thiết bị'), ('vsmt', 'ô tick mục 4'), ('atld', 'ô tick mục 5')) if not vt.get(k)]
    return ngay


def _cum(l, chu):
    """Vi tri cac cum chu `chu` trong dong: [(x0, x1)]."""
    cs = l['chars']; s = ''.join(c['text'] for c in cs); kq = []
    s2 = s.replace(' ', ''); chu2 = chu.replace(' ', '')
    idx = [i for i, c in enumerate(cs) if c['text'] != ' ']
    st = 0
    while True:
        j = s2.find(chu2, st)
        if j < 0: break
        a, b = cs[idx[j]], cs[idx[j + len(chu2) - 1]]
        kq.append((a['x0'], b['x1'])); st = j + len(chu2)
    return kq


def _o_tick(pg, l, so):
    """3 o vuong truoc 'Tốt', 'Bình thường', 'Kém' -> [{trang, x0, top, x1, bottom}] (toa do pdfplumber)."""
    kq = []
    for chu in ('Tốt', 'Bình thường', 'Kém'):
        c = _cum(l, chu)
        if not c: return []
        x0 = c[0][0]
        r = [r for r in pg.rects if r['x1'] <= x0 + 1 and x0 - r['x1'] < 25 and r['top'] < l['bottom'] + 2 and r['bottom'] > l['top'] - 2 and 5 < (r['x1'] - r['x0']) < 20]
        if not r: return []
        r = max(r, key=lambda r: r['x1'])
        kq.append({'trang': so, 'x0': r['x0'], 'top': r['top'], 'x1': r['x1'], 'bottom': r['bottom']})
    return kq


def de_xuat(ngay, dm, tran=None):
    kq = []
    for n in ngay:
        k = tinh_nhan_luc([(x['ten'], x['kl'], x['dv']) for x in n['muc']], dm, tran)
        y, mth, d = n['ngay'].split('-')
        kq.append(dict(ngay=n['ngay'], ngay_vn='%s/%s/%s' % (d, mth, y), so_trang=len(n['trang']), so_viec=len(n['muc']),
                       cong_nhan=k['cong_nhan'], cong=k['cong'], thiet_bi=k['thiet_bi'], thieu=k['thieu'],
                       cong_nhan_dm=k['cong_nhan_dm'], xe_dm=k['xe_dm'], vuot=k['vuot'], chi_cung_cap=k['chi_cung_cap'],
                       doc_loi=n.get('doc_loi', []), thieu_o=n.get('thieu_o', []),
                       sang='', chieu='', mua_sang=False, mua_chieu=False, vsmt='binh_thuong', atld='binh_thuong'))
    return kq


# ---------------- Dien vao PDF ----------------
FONT_SO_TRANG = 'Carlito-Regular.ttf'                  # chu khong chan nhu mau "Trang 34"
_CHAN = re.compile(r'^Ngày\s+\d{1,2}\s+tháng\s+\d{1,2}\s+năm\s+\d{4}\s*-\s*\d+/\d+$')


def dien(du_lieu, gia_tri, thu_muc_font, log=print, giu_ky=False, ngay=None, so_trang=None):
    """gia_tri: {'YYYY-MM-DD': {sang, chieu, mua_sang, mua_chieu, cong_nhan, thiet_bi, vsmt, atld}}
    so_trang: {'bat_dau': 1} -> ghi "Trang N" giua chan moi trang (cung hang chu "Ngày … - k/n"); None: khong danh so."""
    import xuat_pdf, bcsc_sua as B
    ngay = ngay or doc_nhat_ky(du_lieu)
    ttf = xuat_pdf.TTF(open(os.path.join(thu_muc_font, B.FONT_CHU), 'rb').read(), B.FONT_CHU)
    dung = {}

    def hexs(s):
        o = []
        for ch in unicodedata.normalize('NFC', s):
            g = ttf.cmap.get(ord(ch))
            if not g: raise ValueError('font không có chữ "%s" (trong "%s")' % (ch, s))
            dung[g] = ch; o.append('%04X' % g)
        return '<' + ''.join(o) + '>'

    sua = {}

    def chu(so, s, x, y, sz, giua=False):
        if not s: return
        if giua: x -= ttf.rong(s, sz) / 2
        sua.setdefault(so, [None, [], []])[1].append('BT /HCBC %.2f Tf 0 0 0 rg 1 0 0 1 %.2f %.2f Tm %s Tj ET' % (sz, x, y, hexs(s)))

    import pdfplumber
    _pl = pdfplumber.open(io.BytesIO(du_lieu))
    H = {i: float(p.height) for i, p in enumerate(_pl.pages)}
    n_so = 0; fonts_them = []
    if so_trang:
        tts = xuat_pdf.TTF(open(os.path.join(thu_muc_font, FONT_SO_TRANG), 'rb').read(), FONT_SO_TRANG); dung_s = {}
        try: n0 = int(so_trang.get('bat_dau', 1))
        except (TypeError, ValueError): n0 = 1
        for i, pg in enumerate(_pl.pages):
            y = 24.0
            for l in reversed(pg.extract_text_lines()):
                if _CHAN.match(l['text'].strip()): y = _y(l['chars'][0]) + 1.5; break
            s = 'Trang %d' % (n0 + i); o = []
            for ch in s:
                g_ = tts.cmap.get(ord(ch)); dung_s[g_] = ch; o.append('%04X' % g_)
            x = float(pg.width) / 2 - tts.rong(s, 12) / 2
            sua.setdefault(i, [None, [], []])[1].append('BT /HCSO 12 Tf 0 0 0 rg 1 0 0 1 %.2f %.2f Tm <%s> Tj ET' % (x, y, ''.join(o)))
            n_so += 1
        fonts_them.append(('/HCSO', B._font_obj(tts, dung_s, 'HCNKSO+Carlito')))
    n_tick = 0; n_ngay = 0
    for n in ngay:
        g = gia_tri.get(n['ngay'])
        if not g: continue
        n_ngay += 1; vt = n['vi_tri']
        tt, nd, cot = vt.get('thoi_tiet'), vt.get('nhiet_do'), vt.get('cot')
        if cot and nd:
            for cx, buoi in zip(cot, ('sang', 'chieu')):
                chu(nd['trang'], (g.get(buoi) or '').strip(), cx, nd['y'], nd['size'], True)
        if cot and tt and (g.get('mua_sang') or g.get('mua_chieu')):
            if tt['x_xoa'] is not None: sua.setdefault(tt['trang'], [None, [], []])[2].append((tt['x_xoa'], tt['y']))
            for cx, buoi in zip(cot, ('sang', 'chieu')):
                chu(tt['trang'], 'Có mưa' if g.get('mua_' + buoi) else 'Bình thường', cx, tt['y'], tt['size'], True)
        cn = vt.get('cong_nhan')
        try: so_cn = int(g.get('cong_nhan') or 0)
        except (TypeError, ValueError): so_cn = 0
        if cn and so_cn > 0: chu(cn['trang'], '%02d' % so_cn, cn['x'], cn['y'], cn['size'], True)
        tb = vt.get('thiet_bi'); s = re.sub(r'\s+', ' ', str(g.get('thiet_bi') or '')).strip()
        if tb and s:
            sz = tb['size']; w = ttf.rong(s, sz); cho = tb['x_max'] - tb['x']
            if w > cho: sz = max(7.0, sz * cho / w)
            chu(tb['trang'], s, tb['x'], tb['y'], sz)
        for k in ('vsmt', 'atld'):
            o = vt.get(k); ch = g.get(k)
            if not o or ch not in TICK: continue
            r = o[TICK.index(ch)]; h = H[r['trang']]
            x0, x1, y0, y1 = r['x0'], r['x1'], h - r['bottom'], h - r['top']; w_, h_ = x1 - x0, y1 - y0
            sua.setdefault(r['trang'], [None, [], []])[1].append(
                'q 0 0 0 RG 1.4 w 1 J 1 j %.2f %.2f m %.2f %.2f l %.2f %.2f l S Q' % (
                    x0 + 0.18 * w_, y0 + 0.52 * h_, x0 + 0.42 * w_, y0 + 0.2 * h_, x0 + 0.9 * w_, y0 + 0.92 * h_))
            n_tick += 1
    fonts = ([('/HCBC', B._font_obj(ttf, dung))] if dung else []) + fonts_them
    if giu_ky:
        out, bo, (phang, khong_ve) = B._ghi_sach(du_lieu, sua, fonts, None, giu_hinh=True)
        log('Nhật ký: điền %d ngày, tick %d ô; giữ %d hình chữ ký số' % (n_ngay, n_tick, phang))
        return dict(pdf=out, ngay=n_ngay, tick=n_tick, chu_ky_bo=bo, ky_giu=phang, ky_khong_ve=khong_ve, che_do='giu_ky', so_trang=n_so)
    out, bo, _ = B._ghi_sach(du_lieu, sua, fonts, None)
    log('Nhật ký: điền %d ngày, tick %d ô; bỏ %d chữ ký số' % (n_ngay, n_tick, bo))
    return dict(pdf=out, ngay=n_ngay, tick=n_tick, chu_ky_bo=bo, ky_giu=0, ky_khong_ve=0, che_do='bo_ky', so_trang=n_so)
