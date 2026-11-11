import uno,os,sys,json,openpyxl
sys.path.insert(0,'.')
from roll_report import connect,P,col2n,used_rows
src,dst,logsrc=sys.argv[1:4]
# 要対応の科目一覧（AI作業ログ）
wv=openpyxl.load_workbook(logsrc,data_only=True); lg=wv['AI作業ログ']
need={}; on=False
for r in lg.iter_rows(values_only=True):
    if r[0]=='科目': on=True; continue
    if on and r[0]: need[str(r[0]).strip()]=(r[8],r[3],r[7])
d=connect(); doc=d.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(src)),'_blank',0,(P('Hidden',True),))
sh=doc.Sheets.getByName('PL増減分析'); bc,bh=col2n('BC'),col2n('BH')
YEL=0xFFF176; GRN=0xC8E6C9; ny=ng=0
for r in range(5,used_rows(sh)):
    lab=next((sh.getCellByPosition(c,r).getString().strip() for c in range(0,6) if sh.getCellByPosition(c,r).getString().strip()),None)
    if lab not in need: continue
    kind,d1,dc=need[lab]
    for col,diff,th in ((bc,d1,1_000_000),(bh,dc,3_000_000)):
        cell=sh.getCellByPosition(col,r)
        if cell.getString().strip():
            cell.CellBackColor=GRN; ng+=1
        elif isinstance(diff,(int,float)) and abs(diff)>=th:
            cell.CellBackColor=YEL; ny+=1
note=sh.getCellRangeByName('BC2'); note.setString('黄色：コメントの記入が必要（AIでは書けない）／緑：AIが数字から作成したコメント（要確認）'); note.CellBackColor=YEL
doc.calculateAll()
doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(dst)),(P('FilterName','Calc MS Excel 2007 XML'),))
doc.close(True); print('黄色',ny,'緑',ng)
