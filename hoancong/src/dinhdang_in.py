# -*- coding: utf-8 -*-
"""Chinh dinh dang in toan bo file ho so: A4, vua 1 trang ngang (Bia 1x1), vung in den het chu ky,
   lap tieu de bang, noi chieu cao dong chu dai (chi tang, bo dong dem <10pt, bo BIA).  python3 dinhdang_in.py IN.xlsx OUT.xlsx"""
import openpyxl, warnings, sys, re, math; warnings.filterwarnings('ignore')
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
def fmt(src,out):
    print('@@Định dạng in')
    wb=openpyxl.load_workbook(src); log=[]
    nl=nhan_gach_duoi(wb)          # dinh dang nhan TRUOC khi do chieu cao dong (chu dam rong hon)
    import xuat_pdf as XP
    EV=XP.TinhCongThuc(wb); DO=XP.PDF(None); TH=XP._theme(wb); MDW=XP._mdw(wb,DO)
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
        log.append((s.title,s.print_area,ch))
    log.append(('(nhãn gạch dưới)','',nl))
    wb.save(out)
    return log
if __name__=='__main__':
    for l in fmt(sys.argv[1],sys.argv[2]): print('%-28s vùng in %-40s nới %d dòng'%l)
