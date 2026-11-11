import openpyxl,datetime,json,sys
sys.path.insert(0,'.')
from kai_engine import *
from build_materials import sales_tables
k=openpyxl.load_workbook('kai10.xlsx',data_only=True,read_only=True)
prev=[(r[0],r[1],r[2]) for r in k['9月在庫'].iter_rows(min_row=3,values_only=True) if r[0]]
_,sg=load_rows(k['売上']); _,pg=load_rows(k['仕入']); _,kt=load_rows(k['kintone'])
res=compute(prev,sg,pg,kt,datetime.datetime(2026,10,31))
star,first=kintone_cost(k['kintoneコスト']); det=sales_detail(res,star,first,prev)
mi=matome_inputs(det)
pv,dt=sales_tables(res,det)
cum=(53775000+28623500+5655800+7695400+9069300+5739600+4347900+6874000-1524200+10032500+3251100+8695300+11853800)/1000
m=dict(fudosan_k=dict(month=11853.8,cum=cum),matome=mi,
       sales_pivot=[[('' if v is None else v) for v in r] for r in pv],sales_detail=dt)
json.dump(m,open('materials_10.json','w'),ensure_ascii=False,default=str)
print(mi,cum)
