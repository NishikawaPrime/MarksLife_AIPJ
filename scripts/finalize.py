import uno,os,sys,json,openpyxl
sys.path.insert(0,'.')
from roll_report import connect,P,col2n,n2col,used_rows,arr
from openpyxl.utils import column_index_from_string as ci
src,dst,n,mon,genjson,logjson=sys.argv[1:7]; n=int(n)
gen=json.load(open(genjson)); log=json.load(open(logjson))
wv=openpyxl.load_workbook(src,data_only=True)
pl=wv['PL増減分析']; mo=wv['月次推移']
cur=ci('G')+4*(n-1)  # 前期単月列
rows=[]
for r in range(5,pl.max_row+1):
    lab=next((str(pl.cell(r,c).value).strip() for c in range(1,7) if pl.cell(r,c).value),None)
    if not lab: continue
    if '計' in lab or '損益' in lab or '棚卸' in lab or lab.startswith(('仕入','当期商品','売上原価','他勘定','買取再販売上原価','粗利','売上総')): continue
    p0,p1=pl.cell(r,cur).value,pl.cell(r,cur+1).value
    c0,c1=pl.cell(r,ci('BD')).value,pl.cell(r,ci('BE')).value
    if not all(isinstance(x,(int,float)) for x in (p0,p1,c0,c1)): continue
    if abs(p1)<2 and abs(p0)<2: continue  # 比率行はスキップ
    d1=p1-p0; dc=c1-c0
    flag1=abs(d1)>=1_000_000 and (p0==0 or abs(d1/p0)>=0.10)
    flagc=abs(dc)>=3_000_000 and (c0==0 or abs(dc/c0)>=0.10)
    if not (flag1 or flagc): continue
    key_s=f'{lab}|単月'; key_c=f'{lab}|累計'
    kind='自動作成済' if key_s in gen else ('取引先・品目別の明細が必要' if lab not in ('給料手当','雑給','役員報酬','賞与','賞与引当金繰入','支払利息') else '人員数・社内情報が必要')
    prevc=str(pl.cell(r,ci('BC')).value or '')[:200]
    rows.append([lab,p0,p1,d1,(p1/p0 if p0 else None),c0,c1,dc,kind,prevc])
d=connect()
doc=d.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(src)),'_blank',0,(P('Hidden',True),))
S=doc.Sheets; sh=S.getByName('PL増減分析')
bc,bh=col2n('BC'),col2n('BH')
labrow={}
for r in range(0,used_rows(sh)):
    for c in range(0,6):
        s=sh.getCellByPosition(c,r).getString().strip()
        if s:
            labrow.setdefault(s,r); break
# 前月コメントは消して、自動作成分だけ入れる（前月分はログシートに退避）
for r in range(5,used_rows(sh)):
    sh.getCellByPosition(bc,r).setString(''); sh.getCellByPosition(bh,r).setString('')
for k,v in gen.items():
    lab,scope=k.split('|')
    if lab in labrow: sh.getCellByPosition(bc if scope=='単月' else bh,labrow[lab]).setString(v)
# ログシート
if S.hasByName('AI作業ログ'): S.removeByName('AI作業ログ')
S.insertNewByName('AI作業ログ',0); lg=S.getByName('AI作業ログ')
lines=[['【AI作業ログ】'+mon+'版の自動作成内容'],['']]
lines+=[['■ 自動で行ったこと']]+[['・'+x] for x in log['done']]+[['']]
lines+=[['■ 外部から受け取った手入力値（材料として必要）']]+[[f'・{k}：{v}'] for k,v in log['manual'].items()]+[['']]
lines+=[['■ 要確認']]+[['・'+x] for x in log['check']]+[['']]
lg.getCellRangeByPosition(0,0,0,len(lines)-1).setDataArray(arr(lines))
start=len(lines)+1
hdr=[['差異コメント要対応（単月：差100万円以上かつ10%以上／累計：差300万円以上かつ10%以上）','','','','','','','','',''],
     ['科目','前期単月','当期単月','単月差','増加率','前期累計','当期累計','累計差','コメントに必要な情報','（参考）前月版の単月コメント']]
lg.getCellRangeByPosition(0,start,9,start+1).setDataArray(arr(hdr))
if rows: lg.getCellRangeByPosition(0,start+2,9,start+1+len(rows)).setDataArray(arr(rows))
lg.getColumns().getByIndex(0).Width=9000; lg.getColumns().getByIndex(9).Width=20000
for j in range(1,9): lg.getColumns().getByIndex(j).Width=3200
doc.calculateAll()
doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(dst)),(P('FilterName','Calc MS Excel 2007 XML'),))
doc.close(True)
print('rows flagged',len(rows))
