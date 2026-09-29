# -*- coding: utf-8 -*-
"""Tie-out tu ket qua chay refill_main (globals) -> dict de trang web hien thi."""
import re, openpyxl
from collections import Counter
def _cells(wb):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value,str): yield ws.title,c.coordinate,c.value
def tieout(G):
    wb=G['wb']; qlvh=G['qlvh']; boq=G['boq']; bd=G['bd']
    TM=G.get('TM'); TK=G.get('TK'); QM=G.get('QM','7'); QK=G.get('QK','8')
    ntu=len(qlvh); tong=G['TONG']
    ck=[]
    def add(ok,label,detail=''): ck.append({'ok':bool(ok),'label':label,'detail':detail})
    unk=[d['tu'] for d in qlvh+bd if d['phuong']=='(không rõ)']
    add(not unk,'Quy đổi phường theo tủ','Khớp 100% tủ' if not unk else 'Chưa khớp: '+', '.join(sorted(set(unk))[:10]))
    ngay=tong/ntu if ntu else 0
    add(ntu>0 and abs(ngay-round(ngay))<1e-9,'Khối lượng QLVH','%d tủ × %s ngày = %s'%(ntu,('%g'%ngay),'{:,.0f}'.format(tong).replace(',','.')))
    thb=sum(x['th'] or 0 for x in boq.values()); klbd=sum(x['kl'] or 0 for x in bd)
    add(len(boq)==499,'Bảng hạng mục bảo dưỡng','%d hạng mục'%len(boq))
    add(abs(thb-klbd)<1e-6,'Bảo dưỡng: chi tiết khớp tổng hợp','Chi tiết %g = Tổng hợp %g'%(klbd,thb))
    alltxt=set(v for _,_,v in _cells(wb))
    legal=re.compile(r'^\s*-?\s*(Nghị định|Quyết định|Hướng dẫn|Căn cứ|Thông báo|Văn bản số)')
    miss=[t for t in G['PROTECT'] if (legal.match(t) or 'Căn cứ' in t) and t not in alltxt]
    add(not miss,'Căn cứ pháp lý giữ nguyên','%d dòng khớp mẫu'%len([t for t in G['PROTECT'] if legal.match(t) or 'Căn cứ' in t]) if not miss else 'Bị đổi: '+miss[0][:80])
    # phuong cua quan mau (lay tu dong Dia ban cua mau)
    wards=[w.strip() for w in re.findall(r'(?:[Pp]hường|[Xx]ã)\s+([^,.:]+)',G.get('DB_MAU_TXT',''))]
    keep=re.compile(r'(Gói thầu|Dự toán|Địa điểm\s*:?\s*Phường|Cung cấp dịch vụ|Phi tư vấn|Văn bản số|Phường Xóm Chiếu)')
    left=[]
    for sh,co,v in _cells(wb):
        if keep.search(v) or v in G['PROTECT'] and legal.match(v): continue
        if TM and TK and TM!=TK and re.search(r'(tháng|THÁNG)\s+0?%s(\s+(năm|NĂM)|/\d{4})'%TM,v): left.append((sh,co,'tháng %s'%TM))
        if QM!=QK and re.search(r'\bQ%s\b'%QM,v): left.append((sh,co,'Q%s'%QM))
        for w in wards:
            if w and re.search(r'[Pp]hường\s+'+re.escape(w)+r'\b',v): left.append((sh,co,w)); break
    add(not left,'Không còn nội dung của kỳ/quận mẫu','Sạch' if not left else '; '.join('%s!%s (%s)'%x for x in left[:6]))
    qhm=[]
    for h in G['hm_order']: qhm.append([h,G['agg'][h]])
    bhm=Counter()
    for d in bd: bhm[d['hm']]+=d['kl'] or 0
    return {'checks':ck,'dia_ban':G['DBSTR'],'tong_qlvh':tong,'so_tu':ntu,
            'phuong':Counter(d['phuong'] for d in qlvh).most_common(),
            'qlvh_hm':qhm,'bd_hm':[[k,v] for k,v in bhm.most_common()],
            'files':{k:v.split('/')[-1] for k,v in G['detect_files']().items()}}
