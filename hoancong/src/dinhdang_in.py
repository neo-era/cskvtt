# -*- coding: utf-8 -*-
"""Chinh dinh dang in toan bo file ho so: A4, vua 1 trang ngang (Bia 1x1), vung in den het chu ky,
   lap tieu de bang, noi chieu cao dong chu dai theo cach ngat dong cua Excel (chi tang, bo dong dem <10pt; ca BIA de khong cat chu).  python3 dinhdang_in.py IN.xlsx OUT.xlsx"""
import openpyxl, warnings, sys, re, math, copy; warnings.filterwarnings('ignore')
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.cell.cell import MergedCell
TITLES={'4_G201_TTH_PTH_TMY_THU':'$4:$5'}
_NHAN=re.compile(r'(\s*)(Gói thầu|Dự toán mua sắm|Địa điểm)(\s*:\s*)(\S.*)$',re.S)
def nhan_gach_duoi(wb):
    """Nhan 'Gói thầu:', 'Dự toán mua sắm:', 'Địa điểm:' -> dam + gach duoi, phan noi dung chu thuong (theo mau, anh Lam 17/10/2026).
    Chi doi DINH DANG, chu giu nguyen tung ky tu."""
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    n=0
    for s in wb.worksheets:
        for row in s.iter_rows(max_row=min(s.max_row,60)):
            for x in row:
                v=x.value
                if not isinstance(v,str) or isinstance(x,MergedCell): continue
                m=_NHAN.match(v)
                if not m: continue
                f=x.font; col=f.color if (f.color is not None and f.color.type=='rgb') else None
                def F(**k): return InlineFont(rFont=f.name,sz=f.sz,i=f.i,color=col,**k)
                x.value=CellRichText(*( [TextBlock(F(),m.group(1))] if m.group(1) else [] ),
                                     TextBlock(F(b=True,u='single'),m.group(2)),TextBlock(F(b=True),m.group(3)),TextBlock(F(b=False),m.group(4)))
                n+=1
    return n
def bia_dep(ws,XP,EV,DO,TH,MDW):
    """TRANG BIA phu kin kho A4 ngang (anh Lam 17/10/2026): le 1,27 cm, can giua, khung vien doi, logo lon giua trang,
    cac khoi chu can doi. Chi doi DINH DANG/BO CUC, khong doi chu. Mau Bia khac cau truc (khong co 'HỒ SƠ HOÀN CÔNG' o cot A) -> bo qua."""
    from openpyxl.styles import Border, Side, Font, Alignment
    from openpyxl.worksheet.properties import PageSetupProperties
    from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
    from openpyxl.drawing.xdr import XDRPositiveSize2D
    tr=next((r for r in range(1,min(ws.max_row,40)+1) if isinstance(ws.cell(r,1).value,str) and 'HỒ SƠ HOÀN CÔNG' in ws.cell(r,1).value.upper()),None)
    if not tr: return False
    m=re.search(r'(\d+)$',str(ws.print_area or '')); last=int(m.group(1)) if m else ws.max_row
    # --- trang: A4 ngang, le 0,5", can giua 2 chieu, vua 1 trang
    ps=ws.page_setup; ps.orientation='landscape'; ps.paperSize=9
    pm=ws.page_margins; pm.left=pm.right=pm.top=pm.bottom=0.5; pm.header=pm.footer=0.2
    ws.print_options.horizontalCentered=True; ws.print_options.verticalCentered=True
    ws.sheet_properties.pageSetUpPr=PageSetupProperties(fitToPage=True); ps.fitToWidth=1; ps.fitToHeight=1
    W=(841.89-72)*0.985; H=(595.28-72)*0.985
    def chars(pt):                                   # pt -> do rong cot Excel (nguoc cong thuc px)
        px=pt/0.75; return round((px/MDW*256-int(128/MDW))/256,2)
    # cot C hep lam le phai trong khung; o gop A:B -> A:C de chu can giua ca trang
    wa=W*0.19; wc=10.0; ws.column_dimensions['A'].width=chars(wa); ws.column_dimensions['B'].width=chars(W-wa-wc); ws.column_dimensions['C'].width=chars(wc)
    for mr in list(ws.merged_cells.ranges):
        if mr.min_col==1 and mr.max_col==2 and mr.max_row<=last:
            r1,r2=mr.min_row,mr.max_row; ws.unmerge_cells(str(mr)); ws.merge_cells(start_row=r1,start_column=1,end_row=r2,end_column=3)
    ws.print_area='A1:C%d'%last
    # --- chu: co chu va canh
    def F(c,sz,b): x=ws.cell(*c); fo=copy.copy(x.font); fo.sz=sz; fo.b=b; fo.name=fo.name or 'Times New Roman'; x.font=fo
    for r in (1,2):
        if ws.cell(r,1).value: F((r,1),15,True)
    F((tr,1),40,True)
    info=[r for r in range(tr+1,last+1) if not ws.row_dimensions[r].hidden and ws.cell(r,1).value is not None and ws.cell(r,2).value is not None]
    for r in info:
        F((r,1),14,True); F((r,2),14,False)
        a=ws.cell(r,2).alignment; ws.cell(r,2).alignment=Alignment(horizontal='justify',vertical='top',wrap_text=True)
        ws.cell(r,1).alignment=Alignment(horizontal='left',vertical='top',wrap_text=True,indent=1)
    sau=[r for r in range(tr+1,last+1) if not ws.row_dimensions[r].hidden and ws.cell(r,1).value and r not in info]
    for i,r in enumerate(sau): F((r,1),22 if i==0 else 18,True); ws.cell(r,1).alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    # --- chieu cao: do chu thuc te o bo rong moi
    S=XP.Sheet(ws,EV,DO,TH,MDW)
    def cao(r,c):
        runs,so=S.chu(ws.cell(r,c))
        if not runs: return 0
        return sum(l[1] for l in S.ngat(runs,max(1.0,S.rong_o(r,c)-2*XP.PAD)))+4
    hz={}
    for r in (1,2): hz[r]=(cao(r,1) if ws.cell(r,1).value else 0) or 20
    hz[tr]=cao(tr,1)+18
    for r in info: hz[r]=max(cao(r,1),cao(r,2))+8
    for r in sau: hz[r]=cao(r,1)+6
    logo=list(range(3,tr)); khac=[r for r in range(1,last+1) if r not in hz and r not in logo and not ws.row_dimensions[r].hidden]
    con=H-sum(hz.values())
    L=max(110.0,min(170.0,con*0.75)) if logo else 0
    du=max(0.0,con-L)
    for r in logo: ws.row_dimensions[r].height=round(L/len(logo),2)
    dem=[r for r in khac]                              # dong trong: chia phan du lam khoang tho
    for r in hz: ws.row_dimensions[r].height=round(min(409,hz[r]),2)
    if sau: ws.row_dimensions[sau[0]].height=round(hz[sau[0]]+du*0.45,2)
    if info: ws.row_dimensions[info[0]].height=round(hz[info[0]]+du*0.15,2)
    for r in dem: ws.row_dimensions[r].height=round(max(4.0,du*0.40/len(dem)),2)
    if not dem and sau: ws.row_dimensions[sau[-1]].height=round(hz[sau[-1]]+du*0.40,2)
    # --- khung vien doi quanh trang bia, bo vien trong
    d=Side(style='double'); n=Side(style=None)
    for r in range(1,last+1):
        for c in (1,2,3):
            x=ws.cell(r,c)
            x.border=Border(left=d if c==1 else n, right=d if c==3 else n, top=d if r==1 else n, bottom=d if r==last else n)
    # --- logo: giua trang, cao ~85% vung logo, giu ti le
    for im in getattr(ws,'_images',[])[:1]:
        if not logo: break
        w0,h0=(im.width or 100),(im.height or 100); hh=L*0.88; ww=w0*hh/h0
        if ww>W*0.5: ww=W*0.5; hh=h0*ww/w0
        x=(W-ww)/2; col,off=(0,x) if x<wa else (1,x-wa)
        y=(L-hh)/2
        E=12700
        im.anchor=OneCellAnchor(_from=AnchorMarker(col=col,colOff=int(off*E),row=logo[0]-1,rowOff=int(y*E)),ext=XDRPositiveSize2D(int(ww*E),int(hh*E)))
        im.width,im.height=int(ww/0.75),int(hh/0.75)
    return True

def fmt(src,out):
    print('@@Định dạng in')
    wb=openpyxl.load_workbook(src); log=[]
    nl=nhan_gach_duoi(wb); KY=[]; THU=[]; CHAM=[]          # dinh dang nhan TRUOC khi do chieu cao dong (chu dam rong hon)
    import xuat_pdf as XP, anh_mini; anh_mini.cai()      # giu anh (logo) khi chay tren trinh duyet
    EV=XP.TinhCongThuc(wb); EV.tinh_het(); DO=XP.PDF(None); TH=XP._theme(wb); MDW=XP._mdw(wb,DO)
    for s in wb.worksheets:
        if s.sheet_state!='visible': continue
        pa=s.print_area; c1,c2=1,s.max_column
        if pa:
            m=re.search(r"\$?([A-Z]+)\$?\d*:\$?([A-Z]+)\$?\d*",(pa if isinstance(pa,str) else pa[0]).split(',')[0].split('!')[-1])
            if m: c1,c2=CI(m.group(1)),CI(m.group(2))
        last=max([r for r in range(1,s.max_row+1) if any(s.cell(r,c).value not in (None,'') for c in range(c1,c2+1))]+[1])
        s.print_area='%s1:%s%d'%(L(c1),L(c2),last+1)
        ps=s.page_setup; ps.paperSize=9; ps.scale=None
        s.sheet_properties.pageSetUpPr=PageSetupProperties(fitToPage=True)
        ps.fitToWidth=1; ps.fitToHeight=1 if s.title.startswith('BIA') else 0
        if s.title in TITLES and not s.print_title_rows: s.print_title_rows=TITLES[s.title]
        if not s.title.startswith('BIA'):           # SO TRANG theo mau bien ban: 'Trang x / y' goc phai, nghieng
            for hf in (s.oddFooter,s.evenFooter,s.firstFooter):
                hf.right.text='Trang &P / &N'; hf.right.font='Times New Roman,Italic'; hf.right.size=11
        # CHU KHONG XUONG DONG bi cut mep (vd khoi ky, '+ Bên A: TRUNG TÂM ...'): thu co chu vua du (toi thieu 9)
        # DONG CHAM DIEN ('………') dai hon 1 dong: cat bot dau cham cho vua 1 dong
        S0=XP.Sheet(s,EV,DO,TH,MDW)
        for r in S0.rows:
            if r>last: break
            for c in S0.cols:
                a=S0.mgcell.get((r,c))
                if a and a!=(r,c): continue
                x=s.cell(r,c); v=x.value
                if not isinstance(v,str) or not v.strip() or v.startswith('='): continue
                f=DO.font(XP.ho_font(x.font.name),x.font.b,x.font.i); sz=float(x.font.sz or 11)
                if re.fullmatch(r'[\s.…·]+',v) and len(v.strip())>3:
                    w=S0.rong_o(r,c)-2*XP.PAD-1
                    while len(v.strip())>3 and f.rong(v.strip(),sz)>w: v=v.strip()[:-1]
                    if v!=x.value: x.value=v; CHAM.append('%s!%s'%(s.title,x.coordinate))
                    continue
                if x.alignment.wrap_text or x.alignment.horizontal in ('justify','distributed'): continue
                tw=f.rong(v,sz); av=XP.rong_tran(S0,r,c,x.alignment.horizontal)-2*XP.PAD-S0._lui(x)
                if tw>av+0.5 and av>0:
                    nsz=max(9.0,math.floor(sz*av/tw*2)/2)
                    if nsz<sz:
                        fo=copy.copy(x.font); fo.sz=nsz; x.font=fo; THU.append('%s!%s %g→%g'%(s.title,x.coordinate,sz,nsz))
        # NOI CHIEU CAO DONG theo cach ngat dong cua Excel (do chu bang font cung kho Times/Calibri;
        # hieu chinh khop 99,4% voi 1.809 dong Excel tu canh chieu cao trong file mau) -> khong con chu bi cat khi in.
        S=XP.Sheet(s,EV,DO,TH,MDW); ch=0
        for r in S.rows:
            if r>last: break
            cur=s.row_dimensions[r].height
            if cur is not None and cur<10: continue                      # dong dem cua mau
            need=0
            for c in S.cols:
                a=S.mgcell.get((r,c))
                if a and (a!=(r,c) or S.mg[a][0]>r): continue             # o gop nhieu dong: khong tinh
                x=s.cell(r,c)
                if x.value is None or not x.alignment.wrap_text: continue
                runs,so=S.chu(x)
                if so or not runs or not ''.join(t[0] for t in runs).strip(): continue
                LN=S.ngat(runs,max(1.0,S.rong_o(r,c)-2*XP.PAD-S._lui(x)))
                need=max(need,sum(l[1] for l in LN))
            if need>(cur or S.rh.get(r,15.75))+0.5: s.row_dimensions[r].height=math.ceil(need*4)/4; ch+=1
        # KHOI KY nam tron 1 trang (anh Lam 17/10/2026): tinh ngat trang nhu luc in, khoi bi cat -> ngat trang truoc khoi
        for _ in range(6):                      # tieu de bang khong mo coi cuoi trang (lap den khi het)
            S2=XP.Sheet(s,EV,DO,TH,MDW); td=XP.giu_tieu_de(s,S2)
            if not td: break
            KY.append('%s: ngắt trang trước tiêu đề bảng (dòng %d)'%(s.title,td))
        if s.title.startswith('BIA'):
            if bia_dep(s,XP,EV,DO,TH,MDW): KY.append('%s: trình bày lại trang bìa (phủ kín A4, khung đôi, logo giữa)'%s.title)
            log.append((s.title,s.print_area,ch)); continue
        S2=XP.Sheet(s,EV,DO,TH,MDW); kk=XP.giu_khoi_ky(s,S2)
        if kk: KY.append('%s: ngắt trang để khối ký nằm trọn 1 trang (dòng %d)'%(s.title,kk))
        log.append((s.title,s.print_area,ch))
    log.append(('(nhãn gạch dưới)','',nl))
    for x in KY: log.append((x,'',0))
    if THU: log.append(('thu cỡ chữ cho khỏi cụt: '+', '.join(THU),'',len(THU)))
    if CHAM: log.append(('cắt bớt dòng chấm điền: '+', '.join(CHAM),'',len(CHAM)))
    wb.save(out)
    return log
if __name__=='__main__':
    for l in fmt(sys.argv[1],sys.argv[2]): print('%-28s vùng in %-40s nới %d dòng'%l)
