# -*- coding: utf-8 -*-
"""Doc font TrueType nhung trong PDF bao cao su co (Crystal Reports): font subset, ma WinAnsi, KHONG co ToUnicode.
Giai ma: ma byte -> glyph (cmap 3,0: 0xF000+ma) -> chu ky glyph -> ky tu.
 - glyph don: chu ky = md5(byte glyf goc)[:12]  (subset giu nguyen byte glyph cua font goc)
 - glyph ghep (chu co dau): chu goc + cac dau -> ghep chu ky thanh phan, tra bang.
Thuan Python (chay duoc trong Pyodide)."""
import hashlib, struct


class TTFNho:
    def __init__(self, data):
        self.d = d = data
        n = struct.unpack('>H', d[4:6])[0]
        self.bang = {}
        for i in range(n):
            tag, _, off, ln = struct.unpack('>4sIII', d[12 + 16 * i:28 + 16 * i])
            self.bang[tag.decode('latin1')] = (off, ln)
        o = self.bang['head'][0]
        self.upm = struct.unpack('>H', d[o + 18:o + 20])[0]
        self.loca_dai = struct.unpack('>h', d[o + 50:o + 52])[0]
        o = self.bang['maxp'][0]
        self.so_glyph = struct.unpack('>H', d[o + 4:o + 6])[0]
        self._loca()
        self._hmtx()
        self.cmap = self._cmap()

    def _loca(self):
        o, ln = self.bang['loca']; d = self.d; n = self.so_glyph + 1
        if self.loca_dai:
            self.loca = list(struct.unpack('>%dI' % n, d[o:o + 4 * n]))
        else:
            self.loca = [x * 2 for x in struct.unpack('>%dH' % n, d[o:o + 2 * n])]

    def _hmtx(self):
        o = self.bang['hhea'][0]; nh = struct.unpack('>H', self.d[o + 34:o + 36])[0]
        o = self.bang['hmtx'][0]; self.rong = []
        for i in range(self.so_glyph):
            j = min(i, nh - 1); self.rong.append(struct.unpack('>H', self.d[o + 4 * j:o + 4 * j + 2])[0])

    def _cmap(self):
        d = self.d; o = self.bang['cmap'][0]; n = struct.unpack('>H', d[o + 2:o + 4])[0]; kq = {}
        for i in range(n):
            pid, eid, off = struct.unpack('>HHI', d[o + 4 + 8 * i:o + 12 + 8 * i])
            t = o + off; fmt = struct.unpack('>H', d[t:t + 2])[0]
            m = {}
            if fmt == 4:
                sc = struct.unpack('>H', d[t + 6:t + 8])[0] // 2
                ends = struct.unpack('>%dH' % sc, d[t + 14:t + 14 + 2 * sc]); p = t + 16 + 2 * sc
                starts = struct.unpack('>%dH' % sc, d[p:p + 2 * sc]); p += 2 * sc
                deltas = struct.unpack('>%dh' % sc, d[p:p + 2 * sc]); p += 2 * sc
                ro = p; ros = struct.unpack('>%dH' % sc, d[p:p + 2 * sc])
                for k in range(sc):
                    for c in range(starts[k], ends[k] + 1):
                        if c == 0xFFFF: continue
                        if ros[k] == 0:
                            g = (c + deltas[k]) & 0xFFFF
                        else:
                            a = ro + 2 * k + ros[k] + 2 * (c - starts[k])
                            g = struct.unpack('>H', d[a:a + 2])[0]
                            if g: g = (g + deltas[k]) & 0xFFFF
                        if g: m[c] = g
            elif fmt == 0:
                for c in range(256):
                    g = d[t + 6 + c]
                    if g: m[c] = g
            elif fmt == 6:
                fc, cnt = struct.unpack('>HH', d[t + 6:t + 10])
                for k in range(cnt):
                    g = struct.unpack('>H', d[t + 10 + 2 * k:t + 12 + 2 * k])[0]
                    if g: m[fc + k] = g
            if m and (pid, eid) == (3, 0): kq = m          # uu tien bang symbol (3,0)
            elif m and not kq: kq = m
        return kq

    def glyph_cua_ma(self, ma):
        return self.cmap.get(0xF000 + ma) or self.cmap.get(ma)

    def tho(self, g):
        if g is None or g >= self.so_glyph: return b''
        o = self.bang['glyf'][0]
        return self.d[o + self.loca[g]:o + self.loca[g + 1]]

    def thanh_phan(self, g):
        """[(gid, dx, dy)] neu la glyph ghep, None neu glyph don."""
        b = self.tho(g)
        if len(b) < 10 or struct.unpack('>h', b[:2])[0] >= 0: return None
        p = 10; kq = []
        while True:
            fl, gi = struct.unpack('>HH', b[p:p + 4]); p += 4
            if fl & 1: dx, dy = struct.unpack('>hh', b[p:p + 4]); p += 4
            else: dx, dy = struct.unpack('>bb', b[p:p + 2]); p += 2
            if fl & 8: p += 2
            elif fl & 0x40: p += 4
            elif fl & 0x80: p += 8
            kq.append((gi, dx, dy))
            if not fl & 0x20: break
        return kq

    def chu_ky(self, g):
        """Chu ky khong phu thuoc subset: glyph don -> md5 byte; glyph ghep -> 'goc+dau1+dau2'."""
        tp = self.thanh_phan(g)
        if tp is None:
            b = self.tho(g)
            return hashlib.md5(b).hexdigest()[:12] if b else 'rong'
        return '+'.join(self.chu_ky(gi) for gi, _, _ in tp)

    def hinh(self, g):
        """Chu ky theo HINH net (bo lenh hinting): ban font Times New Roman khac nhau thuong chi khac hinting,
        net chu giu nguyen -> van nhan ra. Glyph don: md5(diem cuoi contour + co on-curve + toa do tuyet doi)."""
        b = self.tho(g)
        if not b: return 'rong'
        nc = struct.unpack('>h', b[:2])[0]
        if nc < 0:
            return '+'.join(self.hinh(gi) for gi, _, _ in self.thanh_phan(g))
        try:
            p = 10; ends = struct.unpack('>%dH' % nc, b[p:p + 2 * nc]); p += 2 * nc
            il = struct.unpack('>H', b[p:p + 2])[0]; p += 2 + il                     # bo instructions
            n = (ends[-1] + 1) if nc else 0; fl = []
            while len(fl) < n:
                f = b[p]; p += 1; fl.append(f)
                if f & 8:
                    r = b[p]; p += 1; fl += [f] * r
            fl = fl[:n]; xs = []; v = 0
            for f in fl:
                if f & 2: d = b[p]; p += 1; v += d if f & 16 else -d
                elif not f & 16: v += struct.unpack('>h', b[p:p + 2])[0]; p += 2
                xs.append(v)
            ys = []; v = 0
            for f in fl:
                if f & 4: d = b[p]; p += 1; v += d if f & 32 else -d
                elif not f & 32: v += struct.unpack('>h', b[p:p + 2])[0]; p += 2
                ys.append(v)
            raw = repr((ends, [f & 1 for f in fl], xs, ys)).encode()
            return 'h' + hashlib.md5(raw).hexdigest()[:12]
        except Exception:
            return 'loi'


class GiaiMa:
    """Giai ma 1 font PDF: ma byte -> chuoi unicode ('\\ufffd' neu chua co trong bang)."""
    def __init__(self, du_lieu, ten_font):
        import unicodedata
        import bcsc_bang as BB
        BANG, DAU = BB.BANG, BB.DAU; self._BH = getattr(BB, 'BANG_HINH', {}); self._DH = getattr(BB, 'DAU_HINH', {})
        self.t = TTFNho(du_lieu); self.nhom = 'sans' if 'Arial' in ten_font else 'serif'
        self._B, self._D, self._nfc = BANG, DAU, lambda s: unicodedata.normalize('NFC', s)
        self._cache = {}

    def _tra(self, theo_byte, theo_hinh, g):
        v = theo_byte.get((self.nhom, self.t.chu_ky(g)))
        if v is None: v = theo_hinh.get((self.nhom, self.t.hinh(g)))
        return '\ufffd' if v is None else v

    def ky_tu(self, ma):
        if ma in self._cache: return self._cache[ma]
        t = self.t; g = t.glyph_cua_ma(ma); kq = '�'
        if g is not None:
            tp = t.thanh_phan(g)
            if tp is None:
                kq = self._tra(self._B, self._BH, g)
            else:
                goc = self._tra(self._B, self._BH, tp[0][0])
                if goc == 'ı': goc = 'i'
                dau = [self._tra(self._D, self._DH, gi) for gi, _, _ in tp[1:]]
                kq = self._nfc(goc + ''.join(dau))
        self._cache[ma] = kq
        return kq
