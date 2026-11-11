import openpyxl, datetime, collections, sys
def load_rows(ws, header_row=1):
    it=ws.iter_rows(min_row=header_row,values_only=True)
    hdr=list(next(it)); out=[]
    for r in it:
        if r[0] is None: continue
        out.append(dict(zip(hdr,r)))
    return hdr,out
def num(x):
    return x if isinstance(x,(int,float)) else 0
def compute(prev_inv, sales_gl, purch_gl, kintone, base_date):
    """prev_inv: list of (memo,name,amount). returns dict of tables"""
    sales=collections.OrderedDict(); sales_by_acct=collections.defaultdict(lambda: collections.defaultdict(float))
    for r in sales_gl:
        m=r.get('メモタグ'); 
        if m is None: m='(空白)'
        v=num(r.get('貸方金額'))-num(r.get('借方金額'))
        sales[m]=sales.get(m,0)+v; sales_by_acct[m][r['勘定科目']]+=v
    purch=collections.OrderedDict(); purch_by_acct=collections.defaultdict(lambda: collections.defaultdict(float))
    for r in purch_gl:
        m=r.get('メモタグ')
        if m is None: m='(空白)'
        v=num(r.get('借方金額'))-num(r.get('貸方金額'))
        purch[m]=purch.get(m,0)+v; purch_by_acct[m][r['勘定科目']]+=v
    kt={}
    for r in kintone:
        if r.get('案件ID'): kt.setdefault(r['案件ID'],r)
    def kname(m): return (kt.get(m) or {}).get('案件名') or ''
    def kbuy(m): return (kt.get(m) or {}).get('仕入決済実績日')
    def ksell(m): return (kt.get(m) or {}).get('売却決済実績日')
    prev={m:(n,a) for m,n,a in prev_inv}
    # 前月在庫シート
    prev_tbl=[]
    for m,n,a in prev_inv:
        prev_tbl.append(dict(memo=m,name=kname(m) or n,prev=a,sales=sales.get(m,0),purch=purch.get(m,0),buy=kbuy(m),sell=ksell(m)))
    # 当月仕入シート
    cur_tbl=[]
    for m,p in purch.items():
        pa=prev.get(m,(None,0))[1]
        s=sales.get(m,0); sd=ksell(m)
        keep = s==0 and ((sd is None) or (isinstance(sd,datetime.datetime) and sd>base_date))
        cur_tbl.append(dict(memo=m,name=kname(m),prev=pa,purch=p,buy=kbuy(m),sales=s,sell=sd,inv=(pa+p) if keep else 0))
    # 当月在庫
    inv=[]
    import re
    excluded=[]; alerts_pre=[]
    def prog(m): return (kt.get(m) or {}).get('進捗') or ''
    for r in prev_tbl:
        if r['sales']==0 and r['memo'] not in purch:
            if prog(r['memo'])=='没':
                alerts_pre.append(('要確認：前月在庫だがkintone進捗が「没」（除外するか判断）',r['memo'],r['prev']))
            inv.append(dict(memo=r['memo'],name=prev[r['memo']][0],amount=r['prev'],src='前月繰越'))
    for r in cur_tbl:
        if r['inv']!=0:
            if not re.fullmatch(r'(案件|情報)\d+[①-⑩]?',str(r['memo'])):
                excluded.append(('メモタグが案件/情報形式でない→除外',r['memo'],r['inv'])); continue
            if r['inv']<0:
                excluded.append(('当月在庫がマイナス（取消・修正仕訳）→除外',r['memo'],r['inv'])); continue
            inv.append(dict(memo=r['memo'],name=r['name'] or prev.get(r['memo'],('',))[0],amount=r['inv'],src='当月仕入'))
        elif r['sales']!=0 and r['sell'] is None:
            excluded.append(('当月売上ありだがkintone売却決済日なし→在庫から除外',r['memo'],r['prev']+r['purch']))
    # 検証
    cnt=collections.Counter(x['memo'] for x in inv)
    alerts=list(alerts_pre)
    for x in inv:
        if cnt[x['memo']]>1: alerts.append(('重複',x['memo'],x['amount']))
        if sales.get(x['memo'],0)!=0: alerts.append(('売上計上あるのに在庫残',x['memo'],sales[x['memo']]))
    for m,s in sales.items():
        if s!=0 and ksell(m) is None: alerts.append(('売上あり・kintone売却決済日なし',m,s))
    for r in prev_tbl:
        if r['sales']==0 and r['purch']==0: continue
        if r['sales']==0 and r['memo'] not in purch: pass
    return dict(excluded=excluded,sales=sales,sales_by_acct=sales_by_acct,purch=purch,purch_by_acct=purch_by_acct,prev_tbl=prev_tbl,cur_tbl=cur_tbl,inv=inv,alerts=alerts,kt=kt)

def kintone_cost(ws):
    """kintoneコスト raw: cols E..S (E=案件ID,K=事業利益,L=利益率,N=★,S=実績)"""
    star=collections.defaultdict(float); first={}
    for r in ws.iter_rows(min_row=2,values_only=True):
        rid=r[4]
        if not rid: continue
        if r[13]=='★': star[rid]+=num(r[18])
        first.setdefault(rid,(r[10],r[11]))
    return star,first
def sales_detail(res,star,first,prev_inv):
    prev={m:a for m,n,a in prev_inv}
    out=[]
    for m,s in sorted(res['sales'].items(),key=lambda kv:(0 if str(kv[0]).startswith('案件') else 1,str(kv[0]))):
        if m=='(空白)': continue
        p=prev.get(m,0); pu=res['purch'].get(m,0)
        f=s-p-pu; g=star.get(m,0)
        kp,kr=first.get(m,(None,None))
        kt=res['kt'].get(m) or {}
        b,sd=kt.get('仕入決済実績日'),kt.get('売却決済実績日')
        days=(sd-b).days if b and sd else None
        out.append(dict(memo=m,name=kt.get('案件名',''),sales=s,prev=p,purch=pu,freee_gp=f,star=g,final_gp=f-g,kt_gp=kp,kt_rate=kr,gp_diff=(f-g-kp) if kp is not None else None,buy=b,sell=sd,days=days))
    return out

def matome_inputs(det, stale_days=200):
    same=[d for d in det if d['days']==0]
    norm=[d for d in det if d['days']!=0]
    stale=[d for d in norm if d['days'] is not None and d['days']>=stale_days]
    return dict(同日件数=len(same),通常件数=len(norm),通常売上K=sum(d['sales'] for d in norm)/1000,
                通常粗利K=sum(d['freee_gp'] for d in norm)/1000,
                滞留件数=len(stale),滞留売上K=sum(d['sales'] for d in stale)/1000,滞留粗利K=sum(d['freee_gp'] for d in stale)/1000,
                売上合計=sum(d['sales'] for d in det),freee粗利合計=sum(d['freee_gp'] for d in det))
