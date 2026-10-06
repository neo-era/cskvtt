# -*- coding: utf-8 -*-
"""Dieu phoi 1 lan chay tren trang web (Pyodide).
MODE=new : lap ho so moi tu bo bien ban (refill_main) -> neu co BB12+BB13 thi do phat sinh (phatsinh) -> dinh dang in.
MODE=add : file ho so co san (HS_XLSX) -> neu co BB12+BB13 thi do phat sinh -> dinh dang in.
Ket qua: /work/run/out.xlsx ; tra ve dict R cho giao dien."""
import os, sys, json, runpy
import pdfplumber
RUN = '/work/run'

def kind(p):
    with pdfplumber.open(p) as pdf:
        t = ' '.join((pdf.pages[0].extract_text() or '').split()).upper()
    if 'CHI TIẾT, ĐÃ THỰC HIỆN NGOÀI KẾ HOẠCH' in t: return 'BB13'
    if 'PHÁT SINH NGOÀI KẾ HOẠCH' in t: return 'BB12'
    return None

def _run(path):
    """runpy nhung doi SystemExit (thong bao 'SAI FILE ...' cua script) thanh loi doc duoc."""
    try:
        return runpy.run_path(path)
    except SystemExit as e:
        if e.code in (None, 0): return {}
        raise RuntimeError(str(e.code))

def run():
    E = os.environ; mode = E.get('MODE', 'new'); d = E['SRC_DIR']
    pdfs = sorted(os.path.join(d, f) for f in os.listdir(d) if f.lower().endswith('.pdf'))
    ps = {}
    for p in pdfs:
        k = kind(p)
        if not k: continue
        if k in ps: raise RuntimeError('Có 2 file cùng là biên bản %s: %s và %s' % (k[2:], os.path.basename(ps[k]), os.path.basename(p)))
        ps[k] = p
    if len(ps) == 1:
        co = list(ps)[0]; thieu = 'BB13' if co == 'BB12' else 'BB12'
        raise RuntimeError('Có biên bản %s nhưng thiếu biên bản %s (phát sinh ngoài kế hoạch) — chọn đủ cả 2 file' % (co[2:], thieu[2:]))
    if mode == 'add' and len(ps) < len(pdfs):
        la = [os.path.basename(p) for p in pdfs if p not in ps.values()]
        raise RuntimeError('Chế độ bổ sung chỉ nhận biên bản 12, 13 (phát sinh ngoài kế hoạch). File không phải: ' + ', '.join(la))
    R = {'mode': mode, 'checks': []}
    import refill_lib
    if mode == 'new':
        E['OUT_XLSX'] = RUN + '/a.xlsx'
        G = _run('/app/refill_main.py')
        import report
        R.update(report.tieout(G)); R['thang_mau'] = E.get('THANG_MAU'); R['quan_mau'] = E.get('QUAN_MAU')
        hs = RUN + '/a.xlsx'
    else:
        hs = E['HS_XLSX']
    # bang tra -> ban gon json (luu lai trong trinh duyet)
    if not E['MASTER_XLSX'].endswith('.json'):
        rows = refill_lib.master_rows()
        json.dump([[None if x is None else (x if isinstance(x, (int, float)) else str(x)) for x in r] for r in rows], open(RUN + '/master.json', 'w'), ensure_ascii=False)
        R['master_n'] = len(rows)
    else:
        R['master_n'] = len(refill_lib.master_rows())
    if ps:
        E.update(HS_XLSX=hs, PDF12=ps['BB12'], PDF13=ps['BB13'], OUT_XLSX=RUN + '/b.xlsx', PS_LIB='1')
        P = _run('/app/phatsinh.py')
        it12, it13, LOI, kr = P['it12'], P['it13'], P['LOI'], P['khong_ro']
        nvt = sum(len(x['items']) for x in it13)
        R['checks'].append({'ok': not LOI, 'label': 'Phát sinh ngoài KH: chi tiết (BB13) khớp tổng hợp (BB12)',
                            'detail': ('%d hạng mục, %d vị trí; tổng %s' % (len(it12), nvt, '{:,.3f}'.format(sum(y['th'] or 0 for y in it12)).replace(',', ' ').replace('.', ',').replace(' ', '.'))) if not LOI else '; '.join(LOI[:4])})
        R['checks'].append({'ok': not kr, 'label': 'Phát sinh ngoài KH: phường theo tủ', 'detail': 'Khớp 100% tủ' if not kr else 'Chưa tra được: ' + ', '.join(sorted(set(kr))[:10])})
        TH, NAM = P['TH'], P['NAM']
        if mode == 'new' and TH and str(TH) != str(E.get('THANG')):
            R['checks'].append({'ok': False, 'label': 'Kỳ biên bản phát sinh trùng kỳ hồ sơ', 'detail': 'Biên bản 12/13 là tháng %s/%s, hồ sơ lập tháng %s' % (TH, NAM, E.get('THANG'))})
        R['ps'] = {'sheets': [P['S12'].title, P['S13'].title], 'ky': '%s/%s' % (TH, NAM) if TH else '',
                   'hm': [[y['ten'], y['dvt'], y['th'] or 0] for y in it12], 'n_vitri': nvt,
                   'files': [os.path.basename(ps['BB12']), os.path.basename(ps['BB13'])]}
        hs = RUN + '/b.xlsx'
    import dinhdang_in
    log = dinhdang_in.fmt(hs, RUN + '/out.xlsx')
    R['checks'].append({'ok': True, 'label': 'Định dạng in', 'detail': '%d sheet: A4, vừa 1 trang ngang, vùng in tới hết chữ ký; nới %d dòng chữ dài' % (len(log), sum(x[2] for x in log))})
    return R
