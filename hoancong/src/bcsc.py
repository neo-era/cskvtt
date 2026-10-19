# -*- coding: utf-8 -*-
"""Bao cao su co (BCSC, Crystal Reports): doc chu (font ma hoa rieng) theo dong, co toa do.
Dung: tl = DocBCSC(du_lieu_pdf); for so, dongs in tl.cac_trang(): ..."""
import io, re
from bcsc_font import GiaiMa


def ma_cua(text):
    """pdfminer tra ma byte duoi dang ky tu (cp1252) hoac '(cid:N)'."""
    m = re.match(r'\(cid:(\d+)\)$', text)
    if m: return int(m.group(1))
    if len(text) != 1: return None
    try: return text.encode('cp1252')[0]
    except UnicodeEncodeError:
        o = ord(text); return o if o < 256 else None


def _co_tounicode(f):
    """ToUnicode doc duoc that (BCSC co tham chieu ToUnicode nhung tro toi doi tuong hong)."""
    from pdfminer.pdftypes import resolve1, PDFStream
    try:
        t = resolve1(f.get('ToUnicode'))
        return isinstance(t, PDFStream) and b'beginbf' in t.get_data()
    except Exception:
        return False


class DocBCSC:
    def __init__(self, du_lieu):
        import pdfplumber
        self.du_lieu = du_lieu
        self.pdf = pdfplumber.open(io.BytesIO(du_lieu))
        self._gm = {}; self._trang = {}

    def _giai_ma(self, page, ten):
        k0 = (id(page), ten)
        if k0 not in self._trang: self._trang[k0] = self._tim_font(page, ten)
        return self._trang[k0]

    def _tim_font(self, page, ten):
        from pdfminer.pdftypes import resolve1, PDFObjRef
        gm = None
        for k, v in resolve1(page.page_obj.resources.get('Font', {})).items():
            f = resolve1(v)
            if str(resolve1(f.get('BaseFont'))).strip("/'") != ten: continue
            khoa = v.objid if isinstance(v, PDFObjRef) else (ten, id(f))      # 2 subset cung ten khac noi dung: tach rieng
            if khoa in self._gm: return self._gm[khoa]
            if f.get('Encoding') is not None or _co_tounicode(f):         # font binh thuong: pdfminer doc dung
                self._gm[khoa] = None; break
            fd = resolve1(f.get('FontDescriptor')) or {}
            ff = resolve1(fd.get('FontFile2'))
            if ff is not None: gm = GiaiMa(ff.get_data(), ten)
            self._gm[khoa] = gm
            break
        return gm

    def dong(self, page):
        """Danh sach dong: dict(text, x0, x1, top, bottom, size, font, chars=[(ky_tu, char_pdfplumber)]), tu tren xuong."""
        nhom = {}
        for c in page.chars:
            nhom.setdefault(round(c['bottom'], 1), []).append(c)
        kq = []
        for y in sorted(nhom):
            cs = sorted(nhom[y], key=lambda c: c['x0'])
            ky = []; s = ''; truoc = None
            for c in cs:
                gm = self._giai_ma(page, c['fontname'])
                ma = ma_cua(c['text'])
                ch = gm.ky_tu(ma) if gm and ma is not None else c['text']
                if truoc is not None and c['x0'] - truoc['x1'] > 0.25 * c['size'] and not s.endswith(' '):
                    s += ' '
                s += ch; ky.append((ch, c)); truoc = c
            kq.append(dict(text=s.strip(), x0=cs[0]['x0'], x1=cs[-1]['x1'], top=min(c['top'] for c in cs),
                           bottom=y, size=cs[0]['size'], font=cs[0]['fontname'], chars=ky))
        return kq

    def cac_trang(self):
        for i, p in enumerate(self.pdf.pages):
            yield i, p, self.dong(p)
