"""取締役会資料2枚目：借入の内訳・現預金の増減内訳・在庫粗利
材料：借入金返済明細.xlsx（17期シート）、月次報告のBS増減分析、kintone買取再販台帳CSV"""
import openpyxl, pandas as pd, math
from decimal import Decimal, ROUND_HALF_UP
MONTHS=['11月','12月','1月','2月','3月','4月','5月','6月','7月','8月','9月','10月']
COLS=['PJ','運転','設備','賞与','納税']
def loan_table(path):
    ws=openpyxl.load_workbook(path,data_only=True)['17期']
    rows=list(ws.iter_rows(min_row=1,max_row=60,values_only=True)); rec=[]
    for i,r in enumerate(rows):
        if r[2]=='返済' and r[1] in MONTHS:
            b=rows[i-1]
            rec.append([r[1],'借入']+[b[j] or 0 for j in range(3,8)])
            rec.append([r[1],'返済']+[r[j] or 0 for j in range(3,8)])
    return pd.DataFrame(rec,columns=['月','区分']+COLS)
def tz(x): return int(x/1e6) if x>=0 else -int(-x/1e6)
def oku_round(x): return float(Decimal(str(x/1e8)).quantize(Decimal('0.1'),rounding=ROUND_HALF_UP))
def oku_trunc(x): return math.trunc(x/1e8*10)/10
def so(v): return f'＋{v:.1f}億' if v>=0 else f'△{abs(v):.1f}億'
def sd(v): return f'+{v:,}' if v>=0 else f'△{abs(v):,}'
def sheet_bond(loan_xlsx):
    """借入金返済明細の17期シートに書かれた「社債」の金額（あれば）"""
    rows=list(openpyxl.load_workbook(loan_xlsx,data_only=True)['17期'].iter_rows(values_only=True))
    for i,r in enumerate(rows):
        if '社債' in [x for x in r if isinstance(x,str)] and i+1<len(rows):
            v=[x for x in rows[i+1] if isinstance(x,(int,float))]
            if v: return v[0]
    return None
def cf_lines(loan_xlsx,report,mon,include_bond=True,bond_override=None):
    df=loan_table(loan_xlsx); m=mon if mon.endswith('月') else mon+'月'
    x=df[df.月.isin(MONTHS[:MONTHS.index(m)+1])].groupby('区分')[COLS].sum()
    ws=openpyxl.load_workbook(report,data_only=True)['BS増減分析']; B={};C={}
    for r in range(1,ws.max_row+1):
        l=ws.cell(r,1).value
        if l and str(l).strip() not in B:
            l=str(l).strip(); B[l]=ws.cell(r,2).value or 0; C[l]=ws.cell(r,3).value or 0
    d=lambda k:(C[k] if isinstance(C[k],(int,float)) else 0)-(B[k] if isinstance(B[k],(int,float)) else 0)
    O=COLS[1:]
    bor,rep=x.loc['借入'].sum(),x.loc['返済'].sum()
    pjb,pjr=x.loc['借入','PJ'],x.loc['返済','PJ']; ob,orp=x.loc['借入',O].sum(),x.loc['返済',O].sum()
    bond=d('社債') if bond_override is None else bond_override
    loan=[f'【借入{sd(tz(bor))}、返済{sd(-tz(rep))}',f'◇PJ資金{sd(tz(pjb-pjr))}',f'  借入{sd(tz(pjb))}、返済{sd(-tz(pjr))}',
          f'◇運転資金 {sd(tz(ob-orp))}',f'  借入{sd(tz(ob))}、返済{sd(-tz(orp))}',f'◇社債{sd(tz(bond))}']
    cash=d('現預金'); comps={'新店投資等':-x.loc['返済','運転'],'在庫増加ーPJ借入':(pjb-pjr)-d('在庫'),
        '税金・賞与融資の返済':-x.loc['返済',['設備','賞与','納税']].sum(),'運転資金・社債':ob+(bond if include_bond else 0)}
    cd=oku_trunc(cash); disp={k:oku_round(v) for k,v in comps.items()}
    op=round(cd-sum(disp.values()),1)
    cashl=[f"・現預金{so(cd).replace('＋','＋')}",f'　▸営業活動：{so(op)}']+[f'　▸{k}：{so(v)}' for k,v in disp.items()]
    check=dict(借入金残高_明細=df[df.月.isin(MONTHS[:MONTHS.index(m)+1])].shape and None)
    return loan,cashl,dict(raw={k:round(v/1e6,1) for k,v in comps.items()},cash=round(cash/1e6,1),bs_loan=C['借入金'])
def inv_gp(csv,month_end):
    d=pd.read_csv(csv,encoding='cp932',dtype=str)
    buy=pd.to_datetime(d['仕入決済実績日'],errors='coerce'); sell=pd.to_datetime(d['売却決済実績日'],errors='coerce')
    gp=pd.to_numeric(d['事業利益(成約予定)'],errors='coerce').fillna(0); e=pd.Timestamp(month_end)
    sel=(buy<=e)&(sell.isna()|(sell>e))
    return tz(gp[sel].sum()), int(sel.sum())
if __name__=='__main__':
    import sys
    L='/mnt/user-data/uploads/Downloads/借入金返済明細.xlsx'; K='/mnt/user-data/uploads/Downloads/買取再販台帳_20261020T181406+0900.csv'
    for rep,mon,end,ib in [('n08_4.xlsx','8月','2026-08-31',False),('n09_4.xlsx','9月','2026-09-30',True)]:
        a,b,c=cf_lines(L,rep,mon,ib); print(mon,a,b,c,inv_gp(K,end))
