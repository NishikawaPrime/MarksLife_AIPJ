import uno,os,sys
sys.path.insert(0,'.')
from roll_report import connect,P,used_rows,find_row
src,dst=sys.argv[1],sys.argv[2]
import os as _o
adds=[tuple(x.split(':')) for x in _o.environ.get('PL_ADDS','減損損失:不正関連損失').split(',')]
d=connect(); doc=d.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(src)),'_blank',0,(P('Hidden',True),))
S=doc.Sheets
for new,before in adds:
    for sname in ['組替(PL)当年','組替(PL)前年','月次推移','PL増減分析']:
        sh=S.getByName(sname)
        if find_row(sh,new,(5,)) is not None: continue
        r=find_row(sh,before,(5,))
        if r is None: print('!!',sname,'no',before); continue
        sh.Rows.insertByIndex(r,1)
        sh.copyRange(sh.getCellByPosition(0,r).getCellAddress(), sh.getCellRangeByPosition(0,r+1,200,r+1).getRangeAddress())
        for c in range(0,6):
            cell=sh.getCellByPosition(c,r)
            if cell.getString()==before: cell.setString(new)
        # 手入力の値・コメントは消す
        for c in range(6,200):
            cell=sh.getCellByPosition(c,r)
            if cell.getType().value in ('VALUE','TEXT'): cell.setString('')
        print(sname,f'「{new}」を「{before}」の上に追加 (row {r+1})')
    ms=S.getByName('マスタ'); er=used_rows(ms)
    if find_row(ms,new,(1,)) is None:
        ms.getCellByPosition(1,er).setString(new); ms.getCellByPosition(2,er).setString(new); print('マスタに追加')
doc.calculateAll()
doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(dst)),(P('FilterName','Calc MS Excel 2007 XML'),))
doc.close(True)
