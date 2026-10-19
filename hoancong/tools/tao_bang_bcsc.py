# -*- coding: utf-8 -*-
"""Tao bang chu ky glyph -> ky tu cho bao cao su co (chay OFFLINE, can fontTools + Pillow + numpy + pikepdf).
Dung: python3 tools/tao_bang_bcsc.py BCSC.pdf [BCSC2.pdf ...]  -> ghi src/bcsc_bang.py (gop voi bang cu)
Cach nhan: Liberation Serif/Sans cung do rong chu voi Times New Roman/Arial.
 - glyph don: so voi glyph Liberation cung do rong (anh bitmap, IoU lon nhat)
 - glyph ghep = chu goc + dau: nhan rieng tung dau (keo gian vua khung, so voi dau to hop Liberation); luc chay ghep NFC"""
import io, sys, os, unicodedata, pprint
import numpy as np, pikepdf
from fontTools.ttLib import TTFont
from fontTools.pens.recordingPen import DecomposingRecordingPen
from PIL import Image, ImageDraw, ImageFilter
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))
from bcsc_font import TTFNho

LIB = '/usr/share/fonts/truetype/liberation/'
def lib_cua(ten):
    s = 'Sans' if 'Arial' in ten else 'Serif'
    b, i = 'Bold' in ten, 'Italic' in ten
    kieu = 'BoldItalic' if b and i else 'Bold' if b else 'Italic' if i else 'Regular'
    return LIB + 'Liberation%s-%s.ttf' % (s, kieu)

N = 192
def ve(gs, gn, lo=-600, hi=2200, nghieng=0.0):
    pen = DecomposingRecordingPen(gs); gs[gn].draw(pen)
    sc = N / (hi - lo); m = Image.new('1', (N * 2, N), 0); dr = ImageDraw.Draw(m)
    im = np.zeros((N, N * 2), bool); cur = []
    def P(x, y): return ((x + nghieng * y) * sc + 10, (hi - y) * sc)
    for op, a in pen.value:
        if op == 'moveTo': cur = [P(*a[0])]
        elif op in ('lineTo', 'qCurveTo', 'curveTo'): cur += [P(*p) for p in a if p is not None]
        else:
            if len(cur) > 2:
                t = Image.new('1', (N * 2, N), 0); ImageDraw.Draw(t).polygon(cur, fill=1); im ^= np.array(t)
            cur = []
    return im

def _mo(a, r=2):
    return np.asarray(Image.fromarray(a.astype('uint8') * 255).filter(ImageFilter.GaussianBlur(r)), float)

def giong(a, b):
    """do giong 2 anh: cosine sau khi lam mo nhe (chiu duoc lech net nho giua TNR va Liberation)"""
    x, y = _mo(a), _mo(b); n = np.sqrt((x * x).sum() * (y * y).sum())
    return (x * y).sum() / n if n else 0.0

def iou(a, b):
    u = (a | b).sum(); return (a & b).sum() / u if u else (1.0 if not a.any() and not b.any() else 0)

VIET = ''.join(chr(c) for c in range(0x21, 0x7F)) + 'ĐđƠơƯưı–—‘’“”…°×'
DAU = '\u0300\u0301\u0303\u0309\u0302\u0306\u031b'            # huyen sac nga hoi mu trang moc; nang = nam duoi chan chu

def _da_giac(gs, gn):
    pen = DecomposingRecordingPen(gs); gs[gn].draw(pen); P = []; cur = []
    for op, a in pen.value:
        if op == 'moveTo': cur = [a[0]]
        elif op in ('lineTo', 'qCurveTo', 'curveTo'): cur += [p for p in a if p is not None]
        else:
            if len(cur) > 2: P.append(cur)
            cur = []
    return P

def _gian(P, S=20):
    """anh dau, keo gian vua khung (bo anh huong vi tri / ti le)"""
    xs = [p[0] for q in P for p in q]; ys = [p[1] for q in P for p in q]
    x0, x1, y0, y1 = min(xs), max(xs), min(ys), max(ys); im = np.zeros((S, S), bool)
    for q in P:
        t = Image.new('1', (S, S), 0)
        ImageDraw.Draw(t).polygon([((x - x0) / ((x1 - x0) or 1) * (S - 1), (y1 - y) / ((y1 - y0) or 1) * (S - 1)) for x, y in q], fill=1)
        im ^= np.array(t)
    return im, y0, y1

def _no(m, r):
    o = m.copy()
    for dy in range(-r, r + 1):
        for dx in range(-r, r + 1):
            o |= np.roll(np.roll(m, dy, 0), dx, 1)
    return o

def chon(mi, ds, A, r=3):
    """Do giong cao nhat. Rieng cap 'chu nay = chu kia + 1 net' (D/Đ, d/đ, O/Q, c/e, ./,): xet net them do
    (phan cua mau lon khong trung mau nho sau khi no r px) - anh phu >35% net do thi la chu co net, nguoc lai la chu khong net."""
    diem = sorted(((giong(mi, A(c)), c) for c in ds), reverse=True)
    if len(diem) < 2: return diem[0][1], diem
    mid = _no(mi, r)
    for i in range(1, min(4, len(diem))):
        if diem[0][0] - diem[i][0] > 0.06: break
        a, b = diem[0][1], diem[i][1]
        ua = A(a) & ~_no(A(b), r); ub = A(b) & ~_no(A(a), r)
        lon, nho, u = (a, b, ua) if ua.sum() >= ub.sum() else (b, a, ub)
        if u.sum() < 60 or min(ua.sum(), ub.sum()) > 0.25 * u.sum(): continue      # khong phai cap them-1-net
        chon_ = lon if (u & mid).sum() / u.sum() > 0.35 else nho                    # nguong do tren BCSC: dung 0.5-1.0, sai <0.2
        return chon_, diem
    return diem[0][1], diem

def nhan_font(du_lieu, ten, bang, dau, bao):
    t = TTFNho(du_lieu)
    ft = TTFont(io.BytesIO(du_lieu)); gs = ft.getGlyphSet(); order = ft.getGlyphOrder()
    L = TTFont(lib_cua(ten)); lgs = L.getGlyphSet(); lcm = L.getBestCmap(); lw = L['hmtx'].metrics
    ung = {}
    for ch in VIET:
        g = lcm.get(ord(ch))
        if g: ung.setdefault(lw[g][0], []).append(ch)
    mau_dau = {ch: _gian(_da_giac(lgs, lcm[ord(ch)]))[0] for ch in DAU if ord(ch) in lcm}
    anh_lib = {}
    nghieng = 0.0
    if 'Italic' in ten:                                       # Liberation Italic khac dang TNR Italic (vd chu so): them ban dung nghieng
        import math
        nghieng = math.tan(math.radians(-ft['post'].italicAngle or 16.33))
        L2 = TTFont(lib_cua(ten.replace('Italic', ''))); l2gs = L2.getGlyphSet(); l2cm = L2.getBestCmap()
    anh_bien = {}
    def bien(ch):
        if ch not in anh_bien:
            v = [ve(lgs, lcm[ord(ch)])]
            if nghieng and ord(ch) in l2cm: v.append(ve(l2gs, l2cm[ord(ch)], nghieng=nghieng))
            anh_bien[ch] = v
        return anh_bien[ch]
    nhom = 'sans' if 'Arial' in ten else 'serif'
    def don(g, ma):
        mi0 = ve(gs, order[g]) if t.tho(g) else None
        def A(ch):                                            # ban mau giong anh nhat
            k = (g, ch)
            if k not in anh_lib: anh_lib[k] = max(bien(ch), key=lambda m: giong(mi0, m))
            return anh_lib[k]
        ky = t.chu_ky(g)
        if (nhom, ky) in bang: return
        if not t.tho(g): bang[(nhom, ky)] = ' ' if t.rong[g] else ''; return
        ds = ung.get(t.rong[g]) or [c for v in ung.values() for c in v]
        c, diem = chon(mi0, ds, A)
        bang[(nhom, ky)] = c
        bao.append((ten, '%02x' % ma, c, round(diem[0][0], 3), diem[1][1] if len(diem) > 1 else '', round(diem[1][0], 3) if len(diem) > 1 else 0))
    for ma in range(256):
        g = t.glyph_cua_ma(ma)
        if g is None: continue
        tp = t.thanh_phan(g)
        if tp is None: don(g, ma); continue
        don(tp[0][0], ma)
        for gi, dx, dy in tp[1:]:
            ky = t.chu_ky(gi)
            if (nhom, ky) in dau: continue
            im, y0, y1 = _gian(_da_giac(gs, order[gi]))
            if y1 + dy < 0: dau[(nhom, ky)] = '\u0323'; continue        # nam duoi chan chu: dau nang
            sc = sorted(((iou(im, b), ch) for ch, b in mau_dau.items()), reverse=True)
            dau[(nhom, ky)] = sc[0][1]
            bao.append((ten, 'dau', hex(ord(sc[0][1])), round(sc[0][0], 2), hex(ord(sc[1][1])), round(sc[1][0], 2)))

# Sua tay: glyph may nhan sai, da kiem bang mat tren luoi anh (tools/xem_bang_bcsc.py). (nhom, chu ky) -> ky tu
SUA_TAY = {
    ('serif', 'a430ca4bb04b'): 'c',   # TNR Bold c (may doan e)
    ('serif', '28d4e7fdc063'): 'I',   # TNR Bold I (goc cua Ị; may doan !)
    ('serif', 'c6fc7e67e28d'): '5',   # TNR Bold 5 (may doan $)
    ('serif', '9007380a7aa6'): ',',   # TNR Italic , (may doan .)
    ('serif', 'd3092344c8ea'): '6',   # TNR Italic 6 (may doan o)
}

def main(pdfs):
    out = os.path.join(os.path.dirname(__file__), '..', 'src', 'bcsc_bang.py')
    bang = {}
    if os.path.exists(out):
        ns = {}; exec(open(out, encoding='utf-8').read(), ns); bang = dict(ns['BANG']); dau = dict(ns['DAU'])
    else: dau = {}
    bao = []
    for p in pdfs:
        pdf = pikepdf.open(p); da = set()
        for pg in pdf.pages:
            for k, f in pg.Resources.get('/Font', {}).items():
                if f.objgen in da or str(f.Subtype) != '/TrueType' or '/FontFile2' not in f.FontDescriptor: continue
                da.add(f.objgen)
                nhan_font(f.FontDescriptor.FontFile2.read_bytes(), str(f.BaseFont), bang, dau, bao)
    bang.update(SUA_TAY)
    with open(out, 'w', encoding='utf-8') as fh:
        fh.write('# -*- coding: utf-8 -*-\n# Sinh boi tools/tao_bang_bcsc.py - KHONG sua tay. (nhom font, chu ky glyph) -> ky tu\nBANG = ')
        fh.write(pprint.pformat(bang, width=120)); fh.write('\n# (nhom font, chu ky glyph dau) -> dau to hop (combining)\nDAU = ')
        fh.write(pprint.pformat(dau, width=120)); fh.write('\n')
    for r in bao: print(*r)
    print('tong', len(bang), len(dau))

if __name__ == '__main__':
    main(sys.argv[1:])
