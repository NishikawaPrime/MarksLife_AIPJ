"""売上構成分析：追加する月の材料をまとめて params.json を作る
使い方:
  python3 uriage_prep.py --month 202608=棚卸08.xlsx --month 202609=棚卸09.xlsx \
     --journal 仕訳帳.csv --kintone 買取再販台帳.csv --report AI月次報告(最新月).xlsx --out params.json
- 棚卸：売上PIVOTの下の売上明細（メモタグ・案件名・当月売上・A：freee粗利・仕入決済・売却決済・日数）。数式の値が入っていること（AI作成版はLibreOfficeで再計算してから渡す）
- 仕訳帳：freee汎用形式（cp932）。不動産取得税＝勘定科目「仕入高（租税公課）」かつ品目に「不動産取得税」（確定・概算）
- 台帳：kintone買取再販台帳CSV（買取再販ID・案件ID・案件名・仕入決済実績日・売却決済実績日）
- 月次報告：買取再販まとめシートの当年の同日／通常の売上・総利益（千円）
"""
import argparse, json, re, openpyxl, pandas as pd
ap=argparse.ArgumentParser()
ap.add_argument('--month',action='append',required=True); ap.add_argument('--journal',required=True)
ap.add_argument('--kintone',required=True); ap.add_argument('--report',required=True); ap.add_argument('--out',default='params.json')
a=ap.parse_args()
K=pd.read_csv(a.kintone,encoding='cp932',dtype=str)
K['buy']=pd.to_datetime(K['仕入決済実績日'],errors='coerce'); K['sell']=pd.to_datetime(K['売却決済実績日'],errors='coerce')
def sales_detail(path):
    wb=openpyxl.load_workbook(path,data_only=True,read_only=True)
    r=list(wb['売上PIVOT'].iter_rows(values_only=True)); hi=[i for i,x in enumerate(r) if x and x[0]=='メモタグ'][0]
    t=[x[:24] for x in r[hi+1:] if isinstance(x[0],str) and x[0].startswith(('案件','情報','買取'))]
    return pd.DataFrame({'tag':[x[0] for x in t],'name':[x[1] for x in t],'sales':[x[2] for x in t],'gp':[x[5] for x in t],'days':[x[23] for x in t]})
def cands(tag):
    base=re.sub(r'[①-⑳]','',tag); return K[K.案件ID==base]
def mapid(tag,name,ym):
    if tag.startswith('買取再販'): return tag,name,None
    c=cands(tag)
    if len(c)>1:
        c2=c[c.sell.dt.strftime('%Y%m')==str(ym)]
        if len(c2)==1: c=c2
        else:
            c3=c[c.案件名.str.replace('　',' ')==str(name).replace('　',' ')]
            if len(c3)==1: c=c3
    if len(c)==1:
        x=c.iloc[0]; return x.買取再販ID,(name if isinstance(name,str) else x.案件名),x
    return None,name,None
def matome(path):
    ws=openpyxl.load_workbook(path,data_only=True)['買取再販まとめ']; out={};hdr=None;sec=None
    for r in ws.iter_rows(values_only=True):
        if r[1] and str(r[1]).startswith('【'): sec=r[1]
        if r[3]=='11月': hdr=r
        if sec in ('【売上高(単位:千円)】','【総利益(単位:千円)】') and r[2] in ('同日決済','通常決済'):
            for i,h in enumerate(hdr):
                if isinstance(h,str) and h.endswith('月'): out[(sec[1:4],r[2],h)]=r[i+1]
    return out
M=matome(a.report)
pat=re.compile(r'(案件\d+[①-⑳]?|情報\d+|買取再販\d+)')
# 仕訳帳から不動産取得税
J=pd.read_csv(a.journal,encoding='cp932',dtype=str)
parts=[]
for s,sg in [('借方',1),('貸方',-1)]:
    x=J[(J[f'{s}勘定科目']=='仕入高（租税公課）')&J[f'{s}品目'].fillna('').str.contains('不動産取得税')]
    parts.append(pd.DataFrame({'ym':x['取引日'].str[:7].str.replace('/','').str.replace('-','').astype(int),'amt':pd.to_numeric(x[f'{s}金額'])*sg,'memo':x[f'{s}メモ'].fillna('')}))
X=pd.concat(parts)
# 売上計上日（kintoneに売却決済日がない案件の日数計算用）
SALEDATE={}
for s_ in ('借方','貸方'):
    y=J[J[f'{s_}勘定科目'].isin(['売上高','売上高（土地）'])]
    for dt,mm in zip(pd.to_datetime(y['取引日'].str.replace('/','-')),y[f'{s_}メモ'].fillna('')):
        t=pat.findall(mm) if 'pat' in globals() else []
        for tg in t: SALEDATE[(int(dt.strftime('%Y%m')),tg)]=max(SALEDATE.get((int(dt.strftime('%Y%m')),tg),dt),dt)
pat=re.compile(r'(案件\d+[①-⑳]?|情報\d+|買取再販\d+)')
detail=[];tax=[];totals={};flags=[]
for spec in a.month:
    ym,path=spec.split('=',1); ym=int(ym); mon=f'{ym%100}月'
    d=sales_detail(path); adj_s=adj_g=0.0
    for _,r in d.iterrows():
        aid,name,k=mapid(r.tag,r['name'],ym); days=r.days
        if not isinstance(days,(int,float)) or pd.isna(days):
            sd=SALEDATE.get((ym,r.tag))
            if k is not None and pd.notna(k.buy) and pd.isna(k.sell) and sd is not None:
                days=int((sd-k.buy).days)
                flags.append(f'{ym} {r.tag} {name}：kintoneに売却決済日がないため、freeeの売上計上日{sd:%Y/%m/%d}と仕入決済{k.buy:%Y/%m/%d}から{days}日として集計')
            elif k is not None and pd.notna(k.buy) and pd.notna(k.sell):
                days=int((k.sell-k.buy).days)
                flags.append(f'{ym} {r.tag} {name}：棚卸に日数がなく、kintoneの仕入{k.buy:%Y/%m/%d}・売却{k.sell:%Y/%m/%d}から{days}日として集計')
                if days==0: adj_s+=float(r.sales); adj_g+=float(r.gp)  # 月次報告では通常決済に入っている
            else:
                days=0; flags.append(f'{ym} {r.tag} {name}：日数が分からず同日決済（0日）として仮置き。要確認')
                adj_s+=float(r.sales); adj_g+=float(r.gp)
        if aid is None: flags.append(f'{ym} {r.tag}：買取再販IDが台帳で引けず、メモタグのまま記載'); aid=r.tag
        detail.append(dict(ym=ym,id=aid,name=name,sales=int(r.sales),gp=int(round(r.gp)),days=int(days)))
    x=X[X.ym==ym].copy()
    def tid(m):
        t=pat.findall(m)
        if len(t)!=1: return m or '(メモタグなし)'
        aid,_,_=mapid(t[0],None,ym); return aid or t[0]
    x['aid']=x.memo.map(tid)
    g=x.groupby('aid').amt.sum()
    tax+= [dict(ym=ym,aid=k,amt=float(v)) for k,v in g.items() if v!=0]
    ss,sn=M[('売上高','同日決済',mon)],M[('売上高','通常決済',mon)]; gs,gn=M[('総利益','同日決済',mon)],M[('総利益','通常決済',mon)]
    totals[ym]=dict(sales_same=ss+adj_s/1000,sales_norm=sn-adj_s/1000,gp_same=gs+adj_g/1000,gp_norm=gn-adj_g/1000,fudosan=float(x.amt.sum())/1000)
    dsum=sum(r['sales'] for r in detail if r['ym']==ym)/1000
    if abs(dsum-(ss+sn))>0.01: flags.append(f'{ym}：明細の売上合計{dsum:,.3f}千円が月次報告の{ss+sn:,.3f}千円と不一致。要確認')
    print(ym,'明細',sum(1 for r in detail if r['ym']==ym),'件 売上',round(dsum,3),'千円 / 月次報告',round(ss+sn,3),'/ 不動産取得税',round(x.amt.sum()),'円')
json.dump(dict(detail=detail,tax=tax,totals={str(k):v for k,v in totals.items()},flags=flags),open(a.out,'w'),ensure_ascii=False,indent=1)
for f in flags: print('要確認:',f)
