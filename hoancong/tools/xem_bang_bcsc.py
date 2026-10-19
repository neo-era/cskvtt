# -*- coding: utf-8 -*-
"""Ve luoi kiem tra bang giai ma BCSC: moi o = glyph trong PDF + ky tu giai ma duoc. Dung de kiem bang mat.
Dung: python3 tools/xem_bang_bcsc.py BCSC.pdf THU_MUC_RA   (chay tu thu muc hoancong)"""
import io,sys,pikepdf
sys.path.insert(0,'src')
from bcsc_font import TTFNho
from fontTools.ttLib import TTFont
from fontTools.pens.recordingPen import DecomposingRecordingPen
from PIL import Image, ImageDraw, ImageFont
ns={}; 
FT=ImageFont.truetype('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',26); FS=ImageFont.truetype('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',11)
def ve(gs,gn,N=70,lo=-600,hi=2200):
    pen=DecomposingRecordingPen(gs); gs[gn].draw(pen); sc=N/(hi-lo)
    m=Image.new('L',(N,N),0); cur=[]
    import numpy as np; im=np.zeros((N,N),bool)
    for op,a in pen.value:
        if op=='moveTo': cur=[a[0]]
        elif op in('lineTo','qCurveTo','curveTo'): cur+=[p for p in a if p is not None]
        else:
            if len(cur)>2:
                t=Image.new('1',(N,N),0); ImageDraw.Draw(t).polygon([(x*sc+4,(hi-y)*sc) for x,y in cur],fill=1); im^=np.array(t)
            cur=[]
    return Image.fromarray((~im).astype('uint8')*255)
pdf=pikepdf.open(sys.argv[1]); da=set(); fi=0
for pg in pdf.pages:
  for k,f in pg.Resources.get('/Font',{}).items():
    if f.objgen in da: continue
    da.add(f.objgen); ten=str(f.BaseFont); dat=f.FontDescriptor.FontFile2.read_bytes()
    t=TTFNho(dat); ft=TTFont(io.BytesIO(dat)); gs=ft.getGlyphSet(); od=ft.getGlyphOrder()
    from bcsc_font import GiaiMa; GM=GiaiMa(dat,ten)
    items=[(ma,t.glyph_cua_ma(ma)) for ma in range(256) if t.glyph_cua_ma(ma) is not None]
    C=12; R=(len(items)+C-1)//C; W=Image.new('L',(C*110,R*80),255); d=ImageDraw.Draw(W)
    for i,(ma,g) in enumerate(items):
        x=(i%C)*110; y=(i//C)*80
        W.paste(ve(gs,od[g]),(x,y+8)); d.text((x+72,y+25),GM.ky_tu(ma),font=FT,fill=0); d.text((x,y),'%02x'%ma,font=FS,fill=0)
        d.rectangle([x,y,x+109,y+79],outline=200)
    W.save(sys.argv[2]+'/luoi%d.png'%fi); print(fi,ten,len(items)); fi+=1
