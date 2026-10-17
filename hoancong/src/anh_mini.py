# -*- coding: utf-8 -*-
"""Thay the toi thieu cho Pillow (PIL) de openpyxl GIU ANH (logo trang Bia) khi chay trong trinh duyet (Pyodide khong co Pillow).
openpyxl chi can: Image.open(fp) -> doi tuong co .size, .format, .fp, .close(). Chi doc kich thuoc PNG/JPEG/GIF, khong giai ma anh.
Dung: import anh_mini; anh_mini.cai()   (khong lam gi neu da co Pillow that)."""
import io, struct, sys, types, builtins


class Image:
    def __init__(self, data):
        self.fp = io.BytesIO(data)
        self.format, self.size = _kich_thuoc(data)

    def close(self):
        pass


def _kich_thuoc(d):
    if d[:8] == b'\x89PNG\r\n\x1a\n':
        w, h = struct.unpack('>II', d[16:24]); return 'PNG', (w, h)
    if d[:6] in (b'GIF87a', b'GIF89a'):
        w, h = struct.unpack('<HH', d[6:10]); return 'GIF', (w, h)
    if d[:2] == b'\xff\xd8':
        i = 2
        while i + 9 < len(d):
            if d[i] != 0xFF: i += 1; continue
            m = d[i + 1]; ln = struct.unpack('>H', d[i + 2:i + 4])[0]
            if m in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                h, w = struct.unpack('>HH', d[i + 5:i + 9]); return 'JPEG', (w, h)
            i += 2 + ln
    if d[:2] == b'BM':
        w, h = struct.unpack('<ii', d[18:26]); return 'BMP', (w, abs(h))
    return 'PNG', (100, 100)


def open(fp, mode='r'):
    if isinstance(fp, (str, bytes)) and not (isinstance(fp, bytes) and len(fp) > 260):
        with builtins.open(fp, 'rb') as fh: data = fh.read()
    elif isinstance(fp, (bytes, bytearray)):
        data = bytes(fp)
    else:
        if hasattr(fp, 'seek'): fp.seek(0)          # openpyxl doc lai cung 1 luong o _data()
        data = fp.read()
    return Image(data)


def cai():
    """Cai ban thay the vao openpyxl neu khong co Pillow. Tra ve True neu da cai."""
    try:
        import PIL.Image  # noqa: F401  (co Pillow that: khong can)
        return False
    except Exception:
        pass
    m = types.ModuleType('PIL.Image'); m.Image = Image; m.open = open
    pk = types.ModuleType('PIL'); pk.Image = m
    sys.modules['PIL'] = pk; sys.modules['PIL.Image'] = m
    import openpyxl.drawing.image as di
    import openpyxl.reader.drawings as rd
    di.PILImage = m; rd.PILImage = m
    return True
