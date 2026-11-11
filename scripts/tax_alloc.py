import pandas as pd,json,sys
act=json.loads(sys.argv[1]); cutoff=sys.argv[2]
L=pd.read_pickle('long2.pkl').reset_index(drop=True)
L=L[L.date<=cutoff]
L['istax']=L.acct.isin(['仮払消費税','仮受消費税'])
base=L[~L.istax].copy(); tax=L[L.istax].copy()
base['sgn']=base.amt.apply(lambda x:1 if x>0 else -1); tax['sgn']=tax.amt.apply(lambda x:1 if x>0 else -1)
# 仮払(借方+)は借方の本体行、仮受(貸方-)は貸方の本体行に、同じ取引No・同じ税区分で金額按分
tx=tax.groupby(['no','taxcls','acct','sgn']).amt.sum().reset_index()
b=base[base.taxcls.str.contains('%',na=False)]
b=b.groupby(['no','taxcls','sgn','memo']).amt.sum().reset_index()
bt=b.groupby(['no','taxcls','sgn']).amt.sum().rename('btot').reset_index()
m=tx.merge(b,on=['no','taxcls','sgn'],suffixes=('_t','_b')).merge(bt,on=['no','taxcls','sgn'])
m['alloc']=m.amt_t*m.amt_b/m.btot
unmatched=tx.merge(bt,on=['no','taxcls','sgn'],how='left'); um=unmatched[unmatched.btot.isna()]
res=m.groupby(['memo','acct']).alloc.sum().unstack(fill_value=0)
print('按分できなかった税行',len(um),round(um.amt.sum()))
ok=ng=0
for k,(o,p) in act.items():
    ai_o=-res['仮受消費税'].get(k,0) if '仮受消費税' in res else 0
    ai_p=res['仮払消費税'].get(k,0)
    for lab,a,v in [('仮受',ai_o,o),('仮払',ai_p,p)]:
        if v is None: continue
        d=round(a)-v
        if abs(d)<=1: ok+=1
        else: ng+=1; print(k,lab,'AI',round(a),'実物',v,'差',d)
print('一致',ok,'不一致',ng)
