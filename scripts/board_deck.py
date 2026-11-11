import copy, math, sys, openpyxl
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
sys.path.insert(0,'.')
from board_vals import vals
def tz(x):  # 百万円・ゼロ方向切り捨て
    return int(x/1e6) if x>=0 else -int(-x/1e6)
def fm(v):  # △表記
    return f'△{abs(v):,}' if v<0 else f'{v:,}'
def fd(v): return f'△{abs(v):,}' if v<0 else f'+{v:,}'
def pct(x): return f'{x*100:.1f}%' if x>=0 else f'△{abs(x)*100:.1f}%'
def pdiff(x): 
    s=round(x*100,1)
    return f'+{s:.1f}%' if s>0 else (f'±0.0%' if s==0 else f'△{abs(s):.1f}%')
def set_cell(cell,text,red=False):
    tf=cell.text_frame; ps=tf.paragraphs
    parts=text.split(' (',1)
    parts=[parts[0]]+(['('+parts[1]] if len(parts)>1 else [])
    if len(ps)==1 and len(parts)==2: parts=[text]
    for i,p in enumerate(ps):
        t=parts[i] if i<len(parts) else ''
        runs=p.runs
        if not runs: p.add_run(); runs=p.runs
        runs[0].text=t
        for r in runs[1:]: r._r.getparent().remove(r._r)
        if red: runs[0].font.color.rgb=RGBColor(0xC0,0,0)
def set_lines(shape,lines,red_idx=()):
    tf=shape.text_frame; ps=tf.paragraphs
    tmpl=copy.deepcopy(ps[0]._p)
    for p in ps[1:]: p._p.getparent().remove(p._p)
    first=tf.paragraphs[0]
    for i,line in enumerate(lines):
        if i==0: p=first
        else:
            el=copy.deepcopy(tmpl); tf._txBody.append(el); p=tf.paragraphs[-1]
        runs=p.runs
        if not runs: p.add_run(); runs=p.runs
        runs[0].text=line
        for r in runs[1:]: r._r.getparent().remove(r._r)
        if i in red_idx: runs[0].font.color.rgb=RGBColor(0xC0,0,0)
def build(template,out,report,n,mon,budget,bs):
    V=vals(report,n)
    prs=Presentation(template); s1,s2=prs.slides[0],prs.slides[1]
    for sh in s1.shapes:
        if sh.has_text_frame and '段階損益' in sh.text_frame.text: set_lines(sh,[f'2026年{mon}月 段階損益'])
        if sh.has_table: tbl=sh.table
    keys=['売上高 計','売上総損益金額','販売管理費 計','営業損益金額','経常損益金額','当期純損益金額']
    S=V['売上高 計']
    for i,k in enumerate(keys):
        r=2+i; v=V[k]
        for side,(a0,a1,bk,col0) in {'m':('m0','m1','m',2),'c':('c0','c1','c',7)}.items():
            nob=budget is None
            p,a,b=v[a0],v[a1],(1 if nob else budget[bk][i]*1e6)
            sp,sa,sb=S[a0],S[a1],(1 if nob else budget[bk][0]*1e6)
            if i==0:
                txt=[fm(tz(p)),fm(round(b/1e6)),fm(tz(a)),f'{fd(tz(a-p))} (増加率{round((a/p-1)*100) if p else 0:.0f}%)'.replace('(増加率-','(増加率△'),
                     f'{fd(tz(a-b))} (増加率{round((a/b-1)*100):.0f}%)'.replace('(増加率-','(増加率△')]
            else:
                rp,ra,rb=p/sp,a/sa,b/sb
                txt=[f'{fm(tz(p))} ({pct(rp)})',f'{fm(round(b/1e6))} ({pct(rb)})',f'{fm(tz(a))} ({pct(ra)})',
                     f'{fd(tz(a-p))} ({pdiff(ra-rp)})',f'{fd(tz(a-b))} ({pdiff(ra-rb)})']
            if nob: txt[1]='予算未受領'; txt[4]='―'
            for j,t in enumerate(txt):
                set_cell(tbl.cell(r,col0+j),t,red=(j in (1,4) and (nob or budget.get('estimated'))))
    # slide2
    ch=[s for s in s2.shapes if s.has_chart][0].chart
    cd=CategoryChartData(number_format='#,##0'); cd.categories=['期初','実績','期初','実績']
    cd.add_series('その他資産/純資産',bs['chart']['other']); cd.add_series('固定資産/その他負債',bs['chart']['fixed'])
    cd.add_series('棚卸資産/借入金',bs['chart']['inv']); cd.add_series('現金',bs['chart']['cash']); cd.add_series('合計',bs['chart']['total'])
    ch.replace_data(cd)
    tot=max(bs['chart']['total']); ch.value_axis.maximum_scale=math.ceil(tot*1.065/100)*100
    for sh in s2.shapes:
        if not sh.has_text_frame: continue
        t=sh.text_frame.text
        if '貸借対照表' in t: set_lines(sh,[f'2026年{mon}月 貸借対照表'])
        elif t.startswith('・現預金'): set_lines(sh,bs['cash_lines'],red_idx=bs.get('cash_red',range(1,9)))
        elif '【期首】' in t: set_lines(sh,bs['inv_lines'])
        elif '当期純利益' in t: set_lines(sh,bs['eq_lines'])
        elif t.startswith('◇納税'): set_lines(sh,bs['liab_lines'],red_idx=bs.get('liab_red',()))
        elif t.startswith('未収還付'): set_lines(sh,bs['oth_lines'])
        elif '【借入' in t: set_lines(sh,bs['loan_lines'],red_idx=bs.get('loan_red',range(1,9)))
        elif '在庫粗利' in t: set_lines(sh,bs['kintone_lines'],red_idx=bs.get('kintone_red',(0,)))
        elif '通知預金' in t: set_lines(sh,bs['dep_lines'])
        elif ('敷金' in t) or ('附属' in t): set_lines(sh,bs['fixed_lines'])
    prs.save(out)
