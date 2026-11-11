"""半角カタカナを全角に（コメント等の文字列セルのみ）"""
import uno,os,sys,re,unicodedata
sys.path.insert(0,'.')
from roll_report import connect,P
HW=re.compile(r'[｡-ﾟ]+')
src,dst=sys.argv[1],sys.argv[2]
d=connect(); doc=d.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(src)),'_blank',0,(P('Hidden',True),))
n=0
for i in range(doc.Sheets.Count):
    sh=doc.Sheets.getByIndex(i)
    rs=sh.queryContentCells(4)  # STRING
    for addr in rs.getRangeAddresses():
        for r in range(addr.StartRow,addr.EndRow+1):
            for c in range(addr.StartColumn,addr.EndColumn+1):
                cell=sh.getCellByPosition(c,r); s=cell.getString()
                if HW.search(s):
                    t=HW.sub(lambda m:unicodedata.normalize('NFKC',m.group(0)),s)
                    t=t.replace('スマートテラジー','スマート・ストラテジー')
                    cell.setString(t); n+=1
doc.calculateAll()
doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(dst)),(P('FilterName','Calc MS Excel 2007 XML'),)); doc.close(True)
print('fixed',n)
