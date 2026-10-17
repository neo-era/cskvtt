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
def phuong_cua_quan(q):
    q = (q or '').strip().upper()
    q = q[1:] if q.startswith('Q') and q[1:].isdigit() else q
    return QUAN_PHUONG.get(q.lstrip('0') or q, [])
