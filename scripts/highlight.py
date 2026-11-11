import sys,copy
from pptx import Presentation
from pptx.util import Emu,Pt
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from lxml import etree
YEL='FFF176'
ORDER=['ln','noFill','solidFill','gradFill','blipFill','pattFill','grpFill','effectLst','effectDag','highlight','uLnTx','uLn','uFillTx','uFill','latin','ea','cs','sym','hlinkClick','hlinkMouseOver','rtl','extLst']
def hl_run(r):
    rPr=r._r.get_or_add_rPr()
    for h in rPr.findall(qn('a:highlight')): rPr.remove(h)
    h=etree.SubElement(rPr,qn('a:highlight')); c=etree.SubElement(h,qn('a:srgbClr')); c.set('val',YEL)
    kids=sorted(list(rPr),key=lambda e: ORDER.index(etree.QName(e).localname) if etree.QName(e).localname in ORDER else 99)
    for k in list(rPr): rPr.remove(k)
    for k in kids: rPr.append(k)
def hl_paras(shape,idx=None):
    for i,p in enumerate(shape.text_frame.paragraphs):
        if idx is None or i in idx:
            for r in p.runs: hl_run(r)
def legend(slide,prs,text):
    w=Emu(int(prs.slide_width*0.42)); h=Emu(int(prs.slide_height*0.045))
    tb=slide.shapes.add_textbox(prs.slide_width-w-Emu(int(prs.slide_width*0.02)),Emu(int(prs.slide_height*0.01)),w,h)
    tf=tb.text_frame; tf.word_wrap=True; p=tf.paragraphs[0]; r=p.add_run(); r.text=text
    r.font.size=Pt(10); r.font.bold=True; hl_run(r)
    from pptx.enum.text import PP_ALIGN
    p.alignment=PP_ALIGN.RIGHT
src,dst=sys.argv[1],sys.argv[2]
prs=Presentation(src); s1,s2=prs.slides[0],prs.slides[1]
tbl=[s for s in s1.shapes if s.has_table][0].table
for r in range(2,8):
    for c in (3,6,8,11):
        cell=tbl.cell(r,c); cell.fill.solid(); cell.fill.fore_color.rgb=RGBColor.from_string(YEL)
legend(s1,prs,'黄色：予算データ未受領のため、実物資料の予算を転記（予算差も転記値から計算）')
for sh in s2.shapes:
    if not sh.has_text_frame: continue
    t=sh.text_frame.text
    if '在庫粗利' in t: hl_paras(sh,idx={0})
    elif '要確認' in t: hl_paras(sh,idx=set(range(1,20)))
    elif '【期首】' in t: hl_paras(sh,idx={0,1,2})
legend(s2,prs,'黄色：期首の件数・単価は転記／在庫粗利は10/20出力のkintone台帳で計算（月末時点の出力が必要）')
prs.save(dst)
