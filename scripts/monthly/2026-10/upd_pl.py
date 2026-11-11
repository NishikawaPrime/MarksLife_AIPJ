import uno,os,sys
sys.path.insert(0,'/home/claude/oct/run'); sys.path.insert(0,'/home/claude/oct/c')
from roll_report import connect,P,col2n
from pl_text import make,lab,pl
import openpyxl
src,dst=sys.argv[1:3]
wf=openpyxl.load_workbook(src)['PL増減分析']
YEL,GRN=0xFFF176,0xC8E6C9
todo=[]
for col,single in ((55,True),(60,False)):
    for r in range(5,wf.max_row+1):
        c=wf.cell(r,col)
        if c.fill.fgColor.rgb=='FFFFF176' and not c.value:
            t=make(r,single)
            if t: todo.append((r,col,t))
d=connect(); doc=d.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(src)),'_blank',0,(P('Hidden',True),))
sh=doc.Sheets.getByName('PL増減分析'); ny=ng=0
for r,col,t in todo:
    c=sh.getCellByPosition(col-1,r-1); c.setString(t)
    if '要記入' in t: c.CellBackColor=YEL; ny+=1
    else: c.CellBackColor=GRN; ng+=1
sh.getCellRangeByName('BC2').setString('黄色：内訳・金額はAIが仕訳から記載、理由などは会社で要記入／緑：AIが数字から作成したコメント（要確認）')
b=doc.Sheets.getByName('BS増減分析'); E=4
def rep(r,o,n):
    c=b.getCellByPosition(E,r-1); t=c.getString(); assert o in t,(r,o); c.setString(t.replace(o,n))
rep(156,'（受贈益22,015Kの設備譲渡分か：要確認）','（設備の譲渡を受け受贈益22,015Kを計上）')
rep(155,'+12,374K（取得目的・用途：要記入）','+12,374K（受贈：土地と合わせ受贈益31,840K。贈与元・用途：要記入）')
rep(161,'+19,466K（取得目的・用途：要記入）','+19,466K（受贈：建物と合わせ受贈益31,840K。贈与元・用途：要記入）')
doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(dst)),(P('FilterName','Calc MS Excel 2007 XML'),))
doc.close(True); print('written',len(todo),'yellow',ny,'green',ng)
