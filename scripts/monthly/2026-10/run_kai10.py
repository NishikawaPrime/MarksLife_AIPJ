import sys,datetime,openpyxl
sys.path.insert(0,'.')
from tanaoroshi_builder import build,zen
exec(open('run_kai.py').read().split('for mon,prev,tmpl')[0].split('from openpyxl.styles import PatternFill,Font')[1].join(['from openpyxl.styles import PatternFill,Font','']) if False else '')
from run_kai_hl import highlight
mon,prev=10,9
rw=openpyxl.load_workbook('kai10.xlsx',data_only=True)
pk=openpyxl.load_workbook('kai09.xlsx',data_only=True,read_only=True)
res,det,mi=build('kai09.xlsx','out/2026.10棚卸_AI作成.xlsx',rw,openpyxl.load_workbook('kai10.xlsx',data_only=True,read_only=True)['9月在庫'],mon,prev,datetime.datetime(2026,10,31))
n=highlight('out/2026.10棚卸_AI作成.xlsx',mon,res)
print(mon,'在庫',len(res['inv']),sum(x['amount'] for x in res['inv']),'売上明細',len(det),'要確認',len(res['alerts']),'除外',len(res['excluded']),'在庫ハイライト',n)
for a in res['alerts']: print('  A',a)
for a in res['excluded']: print('  X',a)
import pickle; pickle.dump(dict(res={k:v for k,v in res.items() if k in('inv','alerts','excluded')},det=det,mi=mi),open('kai10.pkl','wb'))
