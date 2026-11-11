"""freee汎用形式の仕訳帳CSV(cp932) → 1行1明細のlong形式（借方＋、貸方－）"""
import pandas as pd,sys
src,dst=sys.argv[1],sys.argv[2]
d=pd.read_csv(src,encoding='cp932',dtype=str)
parts=[]
for s,sg in [('借方',1),('貸方',-1)]:
    x=pd.DataFrame({'no':d['No'],'jid':d['仕訳ID'],'date':d['取引日'].str.replace('-','/'),
        'acct':d[f'{s}勘定科目'],'amt':pd.to_numeric(d[f'{s}金額'],errors='coerce')*sg,'taxcls':d[f'{s}税区分'],
        'partner':d[f'{s}取引先名'],'item':d[f'{s}品目'],'dept':d[f'{s}部門'],'memo':d[f'{s}メモ'].fillna(''),'note':d[f'{s}備考']})
    parts.append(x[x.acct.notna()])
L=pd.concat(parts); L['ym']=L.date.str[:7].str.replace('/','-')
L.to_pickle(dst); print(len(L))
