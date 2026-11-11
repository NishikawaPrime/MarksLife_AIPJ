"""取締役会資料：表の数値が改行されないよう、はみ出すセルだけ数値のフォントを縮める（西川さん指示 2026-11-11）"""
import sys
from pptx import Presentation
from pptx.util import Pt
K=0.68  # 1文字の幅（em比）。PowerPoint表示で18ptの6文字が幅1,023,240EMUで折り返した実例に余裕を持たせた値
def fit(prs, base=18.0, minpt=11.0):
    n=0
    for s in prs.slides:
        for sh in s.shapes:
            if not sh.has_table: continue
            tb=sh.table
            for ci,col in enumerate(tb.columns):
                for ri in range(len(tb.rows)):
                    cell=tb.cell(ri,ci); paras=cell.text_frame.paragraphs
                    if not paras or not paras[0].runs: continue
                    t=paras[0].text.strip()
                    if not t or not any(ch.isdigit() for ch in t): continue
                    sizes=[r.font.size.pt for r in paras[0].runs if r.font.size] or [base]
                    cur=max(sizes)
                    avail=(col.width-cell.margin_left-cell.margin_right)/12700
                    need=len(t)*K*cur
                    if need>avail:
                        new=max(minpt,int(avail/(len(t)*K)*2)/2)
                        for r in paras[0].runs: r.font.size=Pt(new)
                        n+=1
    return n
if __name__=='__main__':
    p=Presentation(sys.argv[1]); print('縮めたセル',fit(p)); p.save(sys.argv[2])
