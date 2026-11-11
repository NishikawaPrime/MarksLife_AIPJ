import openpyxl,sys,re
sys.path.insert(0,'.')
from comments import gen
from openpyxl.utils import column_index_from_string as ci, get_column_letter as gl
f=sys.argv[1]; n=int(sys.argv[2]); mon=sys.argv[3]
wb=openpyxl.load_workbook(f,data_only=True)
mtw=wb['買取再販まとめ']; plw=wb['PL増減分析']
class MT:
    cur_prev=gl(4+3*(n-1)); cur=gl(5+3*(n-1)); cum_prev=gl(ci('AO')+3*(n-2)); cum=gl(ci('AP')+3*(n-2))
    def v(self,col,row): return mtw[f'{col}{row}'].value
rows={}
for r in range(1,plw.max_row+1):
    lab=next((str(plw.cell(r,c).value).strip() for c in range(1,7) if plw.cell(r,c).value),None)
    if lab and lab not in rows: rows[lab]=r
cur_cur=ci('H')+4*(n-1); 
def pl(lab,scope,prev):
    r=rows[lab]
    col=(cur_cur-prev) if scope=='単月' else (ci('BE')-prev)
    return plw.cell(r,col).value or 0
out=gen(MT(),pl,mon)
import json; json.dump({f'{k[0]}|{k[1]}':v for k,v in out.items()},open(sys.argv[4],'w'),ensure_ascii=False,indent=1)
print(len(out),'comments')
