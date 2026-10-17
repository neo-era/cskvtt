# -*- coding: utf-8 -*-
"""Quan cu -> phuong/xa moi (sap xep don vi hanh chinh TP.HCM 2025, NQ 1685/NQ-UBTVQH15).
Dung de tu dien o "Phuong cua quan" (PREFER) theo ma quan; tren web build.sh chep sang index.html."""
QUAN_PHUONG = {
    '1': ['Phường Sài Gòn', 'Phường Tân Định', 'Phường Bến Thành', 'Phường Cầu Ông Lãnh'],
    '3': ['Phường Bàn Cờ', 'Phường Xuân Hòa', 'Phường Nhiêu Lộc'],
    '4': ['Phường Xóm Chiếu', 'Phường Khánh Hội', 'Phường Vĩnh Hội'],
    '5': ['Phường Chợ Quán', 'Phường An Đông', 'Phường Chợ Lớn'],
    '6': ['Phường Bình Tây', 'Phường Bình Tiên', 'Phường Bình Phú', 'Phường Phú Lâm'],
    '7': ['Phường Tân Thuận', 'Phường Phú Thuận', 'Phường Tân Mỹ', 'Phường Tân Hưng'],
    '8': ['Phường Chánh Hưng', 'Phường Phú Định', 'Phường Bình Đông'],
    '10': ['Phường Diên Hồng', 'Phường Vườn Lài', 'Phường Hòa Hưng'],
    '11': ['Phường Hòa Bình', 'Phường Phú Thọ', 'Phường Bình Thới', 'Phường Minh Phụng'],
    '12': ['Phường Đông Hưng Thuận', 'Phường Trung Mỹ Tây', 'Phường Tân Thới Hiệp', 'Phường Thới An', 'Phường An Phú Đông'],
    'BT': ['Phường Gia Định', 'Phường Bình Thạnh', 'Phường Bình Lợi Trung', 'Phường Thạnh Mỹ Tây', 'Phường Bình Quới'],
    'PN': ['Phường Đức Nhuận', 'Phường Cầu Kiệu', 'Phường Phú Nhuận'],
    'TB': ['Phường Tân Sơn Hòa', 'Phường Tân Sơn Nhất', 'Phường Tân Hòa', 'Phường Bảy Hiền', 'Phường Tân Bình', 'Phường Tân Sơn'],
    'TP': ['Phường Tây Thạnh', 'Phường Tân Sơn Nhì', 'Phường Phú Thọ Hòa', 'Phường Tân Phú', 'Phường Phú Thạnh'],
    'GV': ['Phường Hạnh Thông', 'Phường An Nhơn', 'Phường Gò Vấp', 'Phường An Hội Đông', 'Phường Thông Tây Hội', 'Phường An Hội Tây'],
    'BTA': ['Phường Bình Tân', 'Phường Bình Hưng Hòa', 'Phường Bình Trị Đông', 'Phường An Lạc', 'Phường Tân Tạo'],
    'NB': ['Xã Nhà Bè', 'Xã Hiệp Phước'],
    'CG': ['Xã Cần Giờ', 'Xã Bình Khánh', 'Xã An Thới Đông', 'Xã Thạnh An'],
    'BC': ['Xã Vĩnh Lộc', 'Xã Tân Vĩnh Lộc', 'Xã Bình Lợi', 'Xã Tân Nhựt', 'Xã Bình Chánh', 'Xã Hưng Long', 'Xã Bình Hưng'],
}
# Ten quan/huyen (khong dau, chu thuong) -> ma. Nhap "Bình Thạnh", "Quận Bình Thạnh", "Q.BT", "bt" deu ra 'BT'.
TEN_QUAN = {
    'binh thanh': 'BT', 'phu nhuan': 'PN', 'tan binh': 'TB', 'tan phu': 'TP', 'go vap': 'GV',
    'binh tan': 'BTA', 'nha be': 'NB', 'can gio': 'CG', 'binh chanh': 'BC',
}
def _khong_dau(s):
    import unicodedata
    s = unicodedata.normalize('NFD', (s or '').replace('đ', 'd').replace('Đ', 'D'))
    return ' '.join(''.join(c for c in s if unicodedata.category(c) != 'Mn').lower().split())
def quan_key(s):
    """Chuan hoa ma quan de so khop: '10', 'Q10', 'Quận 10', '010' -> '10'; 'Bình Thạnh', 'Quận Bình Thạnh', 'BT' -> 'BT'."""
    import re
    k = _khong_dau(s)
    k = re.sub(r'^(quan|huyen|q|h)\s*\.?\s*', '', k) if not k in TEN_QUAN else k
    if k in TEN_QUAN: return TEN_QUAN[k]
    if k.isdigit(): return k.lstrip('0') or '0'
    return k.upper().replace(' ', '')
def phuong_cua_quan(q):
    return QUAN_PHUONG.get(quan_key(q), [])
