import openpyxl,math
from openpyxl.utils import column_index_from_string as ci
LABELS=['売上高 計','売上総損益金額','販売管理費 計','営業損益金額','経常損益金額','当期純損益金額']
def vals(path,n):
    ws=openpyxl.load_workbook(path,data_only=True)['PL増減分析']
    rows={}
    for r in range(1,ws.max_row+1):
        for c in range(1,7):
            v=ws.cell(r,c).value
            if v and str(v).strip() in LABELS and str(v).strip() not in rows: rows[str(v).strip()]=r
    cur=ci('G')+4*(n-1)
    out={}
    for l in LABELS:
        r=rows[l]
        out[l]=dict(m0=ws.cell(r,cur).value or 0,m1=ws.cell(r,cur+1).value or 0,c0=ws.cell(r,ci('BD')).value or 0,c1=ws.cell(r,ci('BE')).value or 0)
    return out
if __name__=='__main__':
    import sys
    for l,v in vals(sys.argv[1],int(sys.argv[2])).items(): print(l,{k:round(x/1e6,3) for k,x in v.items()})
