import openpyxl,datetime,sys,json
sys.path.insert(0,'.')
from kai_engine import *
def block(ws,c1,c2):
    out=[]
    for r in ws.iter_rows(min_row=1,max_row=ws.max_row,min_col=c1,max_col=c2,values_only=True):
        out.append([ (v.strftime('%Y-%m-%d') if isinstance(v,datetime.datetime) else v) for v in r])
    while out and all(v is None for v in out[-1]): out.pop()
    return out
def sales_tables(res,det):
    accts=['売上高','売上高（土地）']
    pv=[['合計 / 差引額','列ラベル','',''],['行ラベル']+accts+['総計']]
    tot=[0,0]
    for m in sorted(k for k in res['sales'] if k!='(空白)'):
        a=[res['sales_by_acct'][m].get(x,0) for x in accts]
        tot=[tot[0]+a[0],tot[1]+a[1]]
        pv.append([m]+[(v if v else '') for v in a]+[sum(a)])
    pv.append(['総計',tot[0],tot[1],sum(tot)])
    hdr=['メモタグ','案件名','当月売上','前月在庫','当月仕入','A：freee粗利','B：kinton★コスト（管理費除く）','C：手入力調整','D：A-B-C 最終粗利','D/売上 粗利率','E：キントーン粗利','キントーン粗利率','D-E ㋐粗利差異','粗利率差異','①仮受消費税','②仮払消費税','㋑消費税差額','粗利差異（消費税調整後）','粗利率差異（消費税調整後）','リフォーム金額','コメント','仕入決済','売却決済','仕入から売却の日数']
    dt=[hdr]
    for d in det:
        fg=d['final_gp']; kp=d['kt_gp']
        dt.append([d['memo'],d['name'],d['sales'],d['prev'],d['purch'],d['freee_gp'],d['star'],'',fg,
                   (fg/d['sales']*100 if d['sales'] else ''),('' if kp is None else kp),('' if d['kt_rate'] is None else d['kt_rate']),
                   ('' if kp is None else fg-kp),('' if (kp is None or not d['sales'] or d['kt_rate'] in (None,'')) else fg/d['sales']*100-d['kt_rate']),
                   '','','','','','','',
                   (d['buy'].strftime('%Y-%m-%d') if d['buy'] else ''),(d['sell'].strftime('%Y-%m-%d') if d['sell'] else ''),('' if d['days'] is None else d['days'])])
    return pv,dt
if __name__=='__main__':
    rep=openpyxl.load_workbook('rep08.xlsx',data_only=True)
    k8=openpyxl.load_workbook('kai08.xlsx',data_only=True,read_only=True)
    prev=[(r[0],r[1],r[2]) for r in k8['7月在庫'].iter_rows(min_row=3,values_only=True) if r[0]]
    _,sg=load_rows(k8['売上']); _,pg=load_rows(k8['仕入']); _,kt=load_rows(k8['kintone'])
    res=compute(prev,sg,pg,kt,datetime.datetime(2026,8,31))
    star,first=kintone_cost(k8['kintoneコスト'])
    det=sales_detail(res,star,first,prev)
    mi=matome_inputs(det); mi['人員']=151
    pv,dt=sales_tables(res,det)
    m=dict(month_index=10,month_label='8月',prev_month_label='7月',bs_title='2026年-8月期末',
           freee_bs=block(rep['freee貼付(BS)'],3,21), freee_pl=block(rep['freee貼付(PL)当年'],4,21),
           haken=1406937, ma=5583294, fudosan_k=dict(month=8508.3,cum=130288.8),
           matome=mi, sales_pivot=[[('' if v is None else v) for v in r] for r in pv], sales_detail=dt)
    json.dump(m,open('materials_08.json','w'),ensure_ascii=False,default=str)
    print('ok',len(m['freee_bs']),len(m['freee_pl']),mi)
