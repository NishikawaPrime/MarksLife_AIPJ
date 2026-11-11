import pandas as pd,json,sys
L=pd.read_pickle('long.pkl')
L['sgn']=(L.amt>0).map({True:1,False:-1})
def alloc(cut, key='jid'):
    X=L[L.date<=cut]
    T=X[X.acct.isin(['仮払消費税','仮受消費税'])]
    B=X[~X.acct.isin(['仮払消費税','仮受消費税']) & X.taxcls.str.contains('%',na=False)]
    tx=T.groupby([key,'taxcls','acct','sgn']).amt.sum().reset_index()
    b=B.groupby([key,'taxcls','sgn','memo']).amt.sum().reset_index()
    bt=b.groupby([key,'taxcls','sgn']).amt.sum().rename('btot').reset_index()
    m=tx.merge(b,on=[key,'taxcls','sgn'],suffixes=('_t','_b')).merge(bt,on=[key,'taxcls','sgn'])
    m['alloc']=m.amt_t*m.amt_b/m.btot
    um=tx.merge(bt,on=[key,'taxcls','sgn'],how='left'); um=um[um.btot.isna()]
    r=m.groupby(['memo','acct']).alloc.sum().unstack(fill_value=0)
    return r,um,m
if __name__=='__main__':
    act=json.load(open('actual.json'))['08']
    for key in ['jid','no']:
        r,um,m=alloc('2026/08/31',key)
        print('key',key,'按分できない税行',len(um),round(um.amt.sum()))
        ok=0;rows=[]
        for k,v in act.items():
            if v['O'] is None and v['P'] is None: continue
            ao=-r['仮受消費税'].get(k,0); ap=r['仮払消費税'].get(k,0)
            do=round(ao)-(v['O'] or 0); dp=round(ap)-(v['P'] or 0)
            rows.append((k,v['O'],round(ao),do,v['P'],round(ap),dp))
        df=pd.DataFrame(rows,columns=['tag','実物O','AI_O','差O','実物P','AI_P','差P'])
        df['一致']=(df.差O.abs()<=1)&(df.差P.abs()<=1)
        print(df.一致.sum(),'/',len(df))
        if key=='jid': df.to_csv('cmp08.csv',index=False); print(df[~df.一致].to_string())
