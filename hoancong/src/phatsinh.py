# -*- coding: utf-8 -*-
"""Do 2 bien ban PHAT SINH NGOAI KE HOACH vao ho so hoan cong:
   BB12 = 'BB xac nhan KL da thuc hien, phat sinh ngoai KH' (bang 6 cot: STT, hang muc, DVT, TH, XN, ghi chu)
   BB13 = 'BB XNKL chi tiet, da thuc hien ngoai KH' (bang 10 cot: ..., quan, duong, tu, vi tri, TH, XN, ghi chu)
Env: HS_XLSX (file ho so dich), PDF12, PDF13, SRC_XLSM (mau G201 goc - lay style dong), MASTER_XLSX, GOI, PREFER, OUT_XLSX"""
import os, re, sys, copy, warnings, difflib; warnings.filterwarnings('ignore')
import pdfplumber, openpyxl
import refill_lib as RL
from refill_lib import norm, rewrite, look, load_master
RL.GOI=os.environ.get('GOI','G201NA'); RL.PREFER=set(p.strip() for p in os.environ.get('PREFER','').split(';') if p.strip())
HS=os.environ['HS_XLSX']; OUT=os.environ.get('OUT_XLSX',''); SRC=os.environ['SRC_XLSM']
if not OUT or os.path.abspath(OUT)==os.path.abspath(HS): sys.exit('THIẾU OUT_XLSX: phải ghi ra file KHÁC file hồ sơ gốc')
def num(s):
    s=(s or '').strip()
    if not s: return None
    if not re.fullmatch(r'[\d.,]+',s): sys.exit('SAI SỐ LIỆU: ô khối lượng "%s" không phải số — kiểm tra PDF đúng biên bản 12/13 chưa'%s)
    return float(s.replace('.','').replace(',','.'))
def check_title(path,key,ten,cam=None):
    with pdfplumber.open(path) as pdf: t=' '.join((pdf.pages[0].extract_text() or '').split()).upper()
    if cam and cam in t: sys.exit('SAI FILE: %s là biên bản chi tiết (có "%s"), không phải %s — kiểm tra có nộp đảo 2 file không'%(os.path.basename(path),cam,ten))
    if key not in t: sys.exit('SAI FILE: %s không phải %s (không thấy "%s" ở trang 1)'%(os.path.basename(path),ten,key))
def J(s): return ' '.join((s or '').split())
# ---------- doc PDF ----------
def ky_pdf(path):
    with pdfplumber.open(path) as pdf: t=pdf.pages[0].extract_text() or ''
    t=' '.join(t.split())     # ky = thang sau 'KE HOACH' (khong lay ngay ky bien ban o dau trang)
    m=re.search(r'K[ẾE]\s*HO[ẠA]CH\s+TH[ÁA]NG\s+0?(\d{1,2})\s+N[ĂA]M\s+(\d{4})',t,re.I)
    return (int(m.group(1)),int(m.group(2))) if m else (None,None)
def parse12(path):
    out=[]
    with pdfplumber.open(path) as pdf:
        for pg in pdf.pages:
            for tb in pg.extract_tables():
                for r in tb:
                    if len(r)<6: continue
                    st=J(r[0])
                    if st in ('I','II','III','IV'):
                        g=J(r[1])
                        if not any(x.get('grp')==g for x in out): out.append({'grp':g})
                    elif st.isdigit():
                        out.append({'stt':int(st),'ten':J(r[1]),'dvt':J(r[2]),'th':num(r[3]),'xn':num(r[4]),'gc':J(r[5])})
    return out
def parse13(path):
    out=[]; cur=None
    with pdfplumber.open(path) as pdf:
        for pg in pdf.pages:
            for tb in pg.extract_tables():
                for r in tb:
                    if len(r)<10: continue
                    r=[x or '' for x in r]; st=J(r[0])
                    if 'Thực hiện' in r[7] or st=='STT': continue
                    if not st and r[1] and not r[3] and not r[7]:          # dong nhom (ten co the bi cat sang cot 3) HOAC ten hang muc tran dong
                        g=J((r[1]+r[2]).replace('\n',' '))
                        if re.match(r'(c[oô]ng t[aá]c|b[aả]o d[uư][oỡ]ng|[IVX]+[.\s])',g,re.I):
                            if not any(x.get('grp')==g for x in out): out.append({'grp':g})
                        elif cur is not None: cur['ten']=J(cur['ten']+' '+g)
                        continue
                    if st.isdigit():
                        cur={'stt':int(st),'ten':J(r[1]),'dvt':J(r[2]),'items':[]}; out.append(cur)
                        if not r[7]: continue
                    if (r[7] or r[3]) and cur is not None:
                        cur['items'].append({'quan':J(r[3]),'duong':J(r[4]),'tu':J(r[5]),'mota':J(r[6]),'th':num(r[7]),'xn':num(r[8]),'gc':J(r[9])})
    return out
# ---------- ten chuan ----------
tpl=openpyxl.load_workbook(SRC)
_t=tpl['6.HTCT_BD_G201']; TN={}
for r in range(33,_t.max_row+1):
    n=_t.cell(r,2).value
    if isinstance(n,str) and _t.cell(r,4).value: TN[norm(n)]=(n.strip(),_t.cell(r,4).value)
FIX=[(r'\bbong(?= đèn)','bóng'),(r'^Bong$','Bóng'),(r'(?<=nâng )mong\b','móng'),(r'\bmong(?= trụ)','móng'),(r'thương xuyên','thường xuyên')]   # chi sua dung ngu canh (khong dung 'bong troc')
def fixw(s):
    for a,b in FIX: s=re.sub(a,b,s)
    return s
def clean(ten,dvt):
    return TN.get(norm(ten)) or (fixw(ten),fixw(dvt))
print('@@Ghi phát sinh ngoài KH (12, 13)')
check_title(os.environ['PDF12'],'PHÁT SINH NGOÀI KẾ HOẠCH','biên bản 12 (xác nhận KL phát sinh ngoài KH)',cam='CHI TIẾT, ĐÃ THỰC HIỆN')
check_title(os.environ['PDF13'],'CHI TIẾT, ĐÃ THỰC HIỆN NGOÀI KẾ HOẠCH','biên bản 13 (XNKL chi tiết ngoài KH)')
p12=parse12(os.environ['PDF12']); p13=parse13(os.environ['PDF13'])
it12=[x for x in p12 if 'stt' in x]; it13=[x for x in p13 if 'stt' in x]
for x in p12:
    if 'grp' in x: x['grp']=fixw(x['grp'])
    else: x['ten'],x['dvt']=clean(x['ten'],x['dvt'])
LOI=[]
# ghep tung hang muc BB13 voi BB12: cung tong KL, ten giong nhat (BB13 hay bi cat chu)
dung=set()
for x in it13:
    s13=round(sum(i['th'] or 0 for i in x['items']),3)
    cand=[y for y in it12 if y['stt'] not in dung and abs((y['th'] or 0)-s13)<1e-6]
    if not cand: LOI.append('BB13 mục %d "%s": tổng %s không khớp hạng mục nào của BB12'%(x['stt'],x['ten'][:50],s13)); continue
    y=max(cand,key=lambda y: difflib.SequenceMatcher(None,norm(x['ten']),norm(y['ten'])).ratio())
    dung.add(y['stt']); x['ten'],x['dvt']=y['ten'],y['dvt']
for y in it12:
    if y['stt'] not in dung: LOI.append('BB12 mục %d "%s" (%s) không có chi tiết ở BB13'%(y['stt'],y['ten'][:50],y['th']))
for y in it12:
    if y['xn'] is not None and abs((y['th'] or 0)-y['xn'])>1e-6: LOI.append('BB12 mục %d: Thực hiện %s ≠ Xác nhận %s'%(y['stt'],y['th'],y['xn']))
# ---------- tim sheet dich ----------
wb=openpyxl.load_workbook(HS)
def text(ws,rmax=12): return ' '.join(str(c.value) for row in ws.iter_rows(max_row=rmax) for c in row if isinstance(c.value,str)).upper()
L12,L13=[],[]
for ws in wb.worksheets:
    if ws.sheet_state!='visible': continue
    t=text(ws)
    if 'CHI TIẾT' in t and 'NGOÀI KẾ HOẠCH' in t: L13.append(ws)
    elif 'PHÁT SINH NGOÀI KẾ HOẠCH' in t: L12.append(ws)
if len(L12)!=1 or len(L13)!=1:
    sys.exit('KHÔNG XÁC ĐỊNH ĐƯỢC SHEET ĐÍCH: BB12 khớp %s, BB13 khớp %s (cần đúng 1 sheet mỗi loại)'%([w.title for w in L12],[w.title for w in L13]))
S12,S13=L12[0],L13[0]
def find(ws,pred,start=1):
    for r in range(start,ws.max_row+1):
        for c in range(1,14):
            v=ws.cell(r,c).value
            if isinstance(v,str) and pred(v.strip()): return r
def need(r,what,ws):
    if r is None: sys.exit('KHÔNG THẤY MỐC %s ở sheet "%s" — mẫu sheet khác thường, hỏi lại anh'%(what,ws.title))
    return r
def tail_of(ws,first):
    r=need(find(ws,lambda v:'ĐẠI DIỆN' in v,first),'"ĐẠI DIỆN" (khối chữ ký)',ws)
    blank=all(ws.cell(r-1,c).value in (None,'') for c in range(1,14))
    return r-1 if blank and r-1>=first else r
def cap(ws,row,nc): return {c:copy.copy(ws.cell(row,c)._style) for c in range(1,nc+1)}
ATTR=('font','border','fill','number_format','alignment','protection')
def put(ws,row,st):
    for c,v in st.items():
        if isinstance(v,dict):                       # style lay tu workbook KHAC -> gan doi tuong, khong gan chi so _style
            cell=ws.cell(row,c)
            for a in ATTR: setattr(cell,a,copy.copy(v[a]))
        else: ws.cell(row,c)._style=copy.copy(v)
# ---------- sheet BB12 ----------
# FONT + DINH DANG BANG: lay dung dong chuan cua bang cung loai trong mau (Times New Roman 12, so #,##0.000)
NUM='#,##0.000'
def capmap(ws,row,cmap):        # cmap: {cot_dich: cot_mau}
    # KHONG copy _style giua 2 workbook (chi so tro sai bang style) -> lay doi tuong that
    return {d:{a:copy.copy(getattr(ws.cell(row,s_),a)) for a in ATTR} for d,s_ in cmap.items()}
t6=tpl['6.HTCT_BD_G201']; M12={1:1,2:2,3:3,4:4,5:6,6:7,7:8}          # BB12: E=TH<-F, F=XN<-G, G=ghi chu<-H
G12=capmap(t6,33,M12); D12=capmap(t6,34,M12); H12=t6.row_dimensions[34].height
first12=need(find(S12,lambda v:v=='(1)'),'"(1)" (dòng đánh số cột)',S12)+1
plan=[('g',x['grp']) if 'grp' in x else ('d',x) for x in p12]
def w12(ws,fd,pr):
    mg=[]; r=fd
    for t,x in plan:
        if t=='g': put(ws,r,G12); ws.cell(r,2,x)
        else:
            put(ws,r,D12); ws.cell(r,1,x['stt']); ws.cell(r,2,x['ten']); ws.cell(r,4,x['dvt']); ws.cell(r,5,x['th']); ws.cell(r,6,'=E%d'%r)
            if x['gc']: ws.cell(r,7,x['gc'])
        ws.row_dimensions[r].height=H12; mg.append((r,r,2,3)); r+=1
    return mg
n12=rewrite(S12,first12,tail_of(S12,first12),7,len(plan),w12,first12)
for r in range(first12,n12):
    for c in (5,6): S12.cell(r,c).number_format=NUM
# ---------- sheet BB13 ----------
t7=tpl['7_G201_TTH_PTH_TMY_THU']; M13={c:c for c in range(3,10)}; M13.update({10:11,11:12,12:13})   # BB13: J=TH<-K, K=XN<-L, L=ghi chu<-M
G13,HM13,PH13,DT13=[capmap(t7,r,M13) for r in (6,8,9,10)]; HG,HH,HP,HD=[t7.row_dimensions[r].height for r in (6,8,9,10)]
by,bn=load_master(); khong_ro=[]
plan=[]
for x in p13:
    if 'grp' in x: plan.append(('g',fixw(x['grp']))); continue
    plan.append(('hm',x)); grp={}
    for i in x['items']:
        m=look(i['tu'],by,bn)
        if not m: khong_ro.append(i['tu'])
        i['phuong']=m[0] if m else '(không rõ)'; i['tuc']=m[3] if m else i['tu']
        grp.setdefault(i['phuong'],[]).append(i)
    for ph,its in grp.items():
        plan.append(('ph',ph)); plan+= [('dt',i) for i in its]
def sph(p): return re.sub(r'^(Phường|Xã)\s+','',p)
def w13(ws,fd,pr):
    mg=[]; r=fd; hmr=phr=None; phs=[]
    def close_ph(end):
        if phr:
            for L in 'JK': ws[L+str(phr)]='=SUM(%s%d:%s%d)'%(L,phr+1,L,end)
    def close_hm():
        if hmr and phs:
            for L in 'JK': ws[L+str(hmr)]='='+'+'.join(L+str(p) for p in phs)
    for t,x in plan:
        if t in ('g','hm'): close_ph(r-1); phr=None; close_hm(); hmr=None; phs=[]
        if t=='g': put(ws,r,G13); ws.row_dimensions[r].height=HG; ws.cell(r,4,x); mg.append((r,r,4,9))
        elif t=='hm': put(ws,r,HM13); ws.row_dimensions[r].height=HH; hmr=r; ws.cell(r,3,x['stt']); ws.cell(r,4,x['ten']); ws.cell(r,5,x['dvt'])
        elif t=='ph': close_ph(r-1); put(ws,r,PH13); ws.row_dimensions[r].height=HP; phr=r; phs.append(r); ws.cell(r,6,'* '+x); mg.append((r,r,6,9))
        else:
            put(ws,r,DT13); ws.row_dimensions[r].height=HD
            ws.cell(r,6,sph(x['phuong'])); ws.cell(r,7,x['duong']); ws.cell(r,8,x['tuc']); ws.cell(r,9,x['mota'])
            ws.cell(r,10,round(x['th'] or 0,3)); ws.cell(r,11,round(x['xn'] or 0,3))
            if x['gc']: ws.cell(r,12,x['gc'])
        r+=1
    close_ph(r-1); close_hm(); return mg
first13=need(find(S13,lambda v:v.lower()=='stt'),'"Stt" (tiêu đề bảng)',S13)+2
n13=rewrite(S13,first13,tail_of(S13,first13),12,len(plan),w13,first13)
for r in range(first13,n13):
    for c in (10,11): S13.cell(r,c).number_format=NUM
# ---------- ky (thang) theo bien ban ----------
TH,NAM=ky_pdf(os.environ['PDF12'])
if not TH: LOI.append('Không đọc được kỳ "... KẾ HOẠCH THÁNG x NĂM y" ở BB12 — tháng trong 2 sheet CHƯA đổi')
# CHI doi o danh sach trang: tieu de (co 'KẾ HOẠCH'), dau cot 'Thực hiện/Xác nhận ...', dong '(Đính kèm ...'. Can cu / ngay ky khong dung.
OKKY=re.compile(r'(KẾ HOẠCH|kế hoạch|^\s*\(?Đính kèm|^\s*(Thực hiện|Xác nhận|Kế hoạch)\b)')
if TH:
    for ws in (S12,S13):
        for row in ws.iter_rows(max_row=first13+2 if ws is S13 else first12):
            for c in row:
                v=c.value
                if not isinstance(v,str): continue
                if 'Căn cứ' in v or v.strip().startswith('-') or 'ngày' in v.lower() or not OKKY.search(v): continue   # khong dung can cu / ngay ky
                nv=re.sub(r'((?:tháng|THÁNG)\s+)0?\d{1,2}(\s+(?:năm|NĂM)\s+)\d{4}',lambda m:m.group(1)+('%02d'%TH if m.group(1).isupper() else str(TH))+m.group(2)+str(NAM),v)
                nv=re.sub(r'(tháng\s+)0?\d{1,2}/\d{4}',lambda m:m.group(1)+'%d/%d'%(TH,NAM),nv)
                if nv!=v: c.value=nv
wb.calculation.fullCalcOnLoad=True
wb.save(OUT)
print('OK: BB12 %d hạng mục, BB13 %d hạng mục / %d vị trí; kỳ %s/%s; sheet "%s" + "%s"'%(len(it12),len(it13),sum(len(x['items']) for x in it13),TH,NAM,S12.title,S13.title))
if khong_ro: print('TỦ KHÔNG TRA ĐƯỢC PHƯỜNG:',sorted(set(khong_ro)))
for l in LOI: print('CẢNH BÁO:',l)
if os.environ.get('PS_LIB')!='1': sys.exit(1 if (LOI or khong_ro) else 0)   # PS_LIB=1: goi tu trang web, doc LOI/khong_ro trong globals
