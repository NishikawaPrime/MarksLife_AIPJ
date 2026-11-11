"""列のグループ化を当月用に組み直す（当月の列だけ見える形）。行のグループはそのまま"""
import uno,os,sys
sys.path.insert(0,'.')
from roll_report import connect,P,col2n
from com.sun.star.table import CellRangeAddress
from com.sun.star.table.TableOrientation import COLUMNS
src,dst,n=sys.argv[1],sys.argv[2],int(sys.argv[3])   # n: 1=11月 … 12=10月
d=connect(); doc=d.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(src)),'_blank',0,(P('Hidden',True),))
S=doc.Sheets
def ra(sh,c1,c2):
    a=CellRangeAddress(); a.Sheet=[S.getByIndex(i).Name for i in range(S.Count)].index(sh.Name); a.StartColumn=c1; a.EndColumn=c2; a.StartRow=0; a.EndRow=0; return a
def reset(sh,c1,c2):
    for _ in range(4):
        try: sh.ungroup(ra(sh,c1,c2),COLUMNS)
        except Exception: pass
    for c in range(c1,c2+1): sh.Columns.getByIndex(c).IsVisible=True
def hide(sh,c1,c2):
    if c1>c2: return
    sh.group(ra(sh,c1,c2),COLUMNS); sh.hideDetail(ra(sh,c1,c2))
# 月次推移：G〜I（前年）＋当月より後の当年の月
mo=S.getByName('月次推移'); reset(mo,col2n('G'),col2n('U'))
hide(mo,col2n('G'),col2n('I'))
if n<12: hide(mo,col2n('J')+n,col2n('U'))
# PL増減分析：F〜（当月ブロックの前）を隠し、当月ブロック（4列）を表示、当月の後〜BBを隠す
pl=S.getByName('PL増減分析'); reset(pl,col2n('F'),col2n('BB'))
cur=col2n('G')+4*(n-1)
hide(pl,col2n('F'),cur-1); hide(pl,cur+4,col2n('BB'))
# 買取再販まとめ：単月（D〜AM、3列ずつ）と累計（AO〜BU、3列ずつ、11-12月から）
mt=S.getByName('買取再販まとめ'); reset(mt,col2n('D'),col2n('AM')); reset(mt,col2n('AO'),col2n('BU'))
m0=col2n('D')+3*(n-1); hide(mt,col2n('D'),m0-1); hide(mt,m0+3,col2n('AM'))
if n>=2:
    k0=col2n('AO')+3*(n-2); hide(mt,col2n('AO'),k0-1); hide(mt,k0+3,col2n('BU'))
else: hide(mt,col2n('AO'),col2n('BU'))
doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(dst)),(P('FilterName','Calc MS Excel 2007 XML'),))
doc.close(True); print('ok')
