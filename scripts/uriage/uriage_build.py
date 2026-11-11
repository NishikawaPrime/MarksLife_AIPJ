"""売上構成分析：params.json（uriage_prep.pyの出力）の月を追加する。LibreOffice(UNO, port 2002)で編集し、テキストボックス等を保つ
使い方: python3 uriage_build.py 前月版.xlsx 出力.xlsx params.json
"""
import uno, pandas as pd, re, sys, json, os
from com.sun.star.beans import PropertyValue
def P(n,v): p=PropertyValue(); p.Name=n; p.Value=v; return p
src,dst,pj=sys.argv[1:4]
prm=json.load(open(pj)); D=pd.DataFrame(prm['detail']); T=pd.DataFrame(prm['tax']); TOT={int(k):v for k,v in prm['totals'].items()}
ctx=uno.getComponentContext()
r=ctx.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',ctx)
c=r.resolve('uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext')
desk=c.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop',c)
doc=desk.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(src)),'_blank',0,(P('Hidden',True),))
S=doc.Sheets
def colidx(s):
    n=0
    for ch in s: n=n*26+ord(ch)-64
    return n-1
def colname(i):
    s='';i+=1
    while i: i,rem=divmod(i-1,26); s=chr(65+rem)+s
    return s
def last_row(sh,col=0):
    cur=sh.createCursor(); cur.gotoEndOfUsedArea(False); r=cur.RangeAddress.EndRow
    while r>0 and sh.getCellByPosition(col,r).getString()=='': r-=1
    return r
months=sorted(TOT)
# 既に入っている月は追加しない（二重追加の防止）
sh17=S.getByName('17期_買取再販明細')
have={int(sh17.getCellByPosition(0,i).getValue()) for i in range(1,last_row(sh17)+1)}
dup=[m for m in months if m in have]
assert not dup, f'既に入っている月があります: {dup}'
def append_detail(name):
    sh=S.getByName(name); lr=last_row(sh)
    tmpl=[sh.getCellByPosition(ci,lr).getFormula() for ci in range(11)]; rn=lr+1
    for i,(_,x) in enumerate(D.iterrows()):
        r0=lr+1+i; rr=r0+1
        sh.copyRange(sh.getCellByPosition(0,r0).CellAddress, sh.getCellRangeByPosition(0,lr,11,lr).RangeAddress)
        sh.getCellByPosition(0,r0).setValue(int(x.ym)); sh.getCellByPosition(1,r0).setString(x.id)
        sh.getCellByPosition(2,r0).setString(str(x['name'])); sh.getCellByPosition(3,r0).setValue(int(x.sales))
        sh.getCellByPosition(4,r0).setValue(int(x.gp)); sh.getCellByPosition(7,r0).setValue(int(x.days))
        for ci in (5,6,8,9,10):
            sh.getCellByPosition(ci,r0).setFormula(re.sub(r'(?<![A-Za-z$])([A-K])%d(?!\d)'%rn, lambda m: m.group(1)+str(rr), tmpl[ci]))
    print(name,'+',len(D),'行')
for n in ['17期_買取再販明細','買取再販明細']: append_detail(n)
sh=S.getByName('不動産取得税'); lr=last_row(sh)
for i,(_,x) in enumerate(T.iterrows()):
    r0=lr+1+i
    sh.getCellByPosition(0,r0).setValue(int(x.ym)); sh.getCellByPosition(1,r0).setValue(float(x.amt))
    sh.getCellByPosition(2,r0).setString('不動産取得税'); sh.getCellByPosition(3,r0).setString(x.aid)
print('不動産取得税 +',len(T),'行')
# 買取再販粗利（17期の表：21行目が年月の見出し）
g=S.getByName('買取再販粗利')
hdr={int(g.getCellByPosition(ci,20).getValue()):ci for ci in range(3,16) if g.getCellByPosition(ci,20).getValue()}
for ym,t in TOT.items():
    ci=hdr[ym]
    for row,v in zip((21,22,29,30,31),(t['sales_same'],t['sales_norm'],t['gp_same'],t['gp_norm'],t['fudosan'])): g.getCellByPosition(ci,row).setValue(v)
# まとめ：月のブロック（2行目に年月）
m=S.getByName('まとめ')
blocks={}
for ci in range(0,200):
    v=m.getCellByPosition(ci,1)
    if v.getValue() and 202500<v.getValue()<203000: blocks[int(v.getValue())]=ci
for ym in months:
    b=blocks[ym]
    m.getCellByPosition(b,24).setFormula('=SUM(%s18:%s23)'%(colname(b),colname(b)))      # 同日列＝明細の合計
    m.getCellByPosition(b+4,24).setFormula('=%s25-SUM(%s25:%s25)'%(colname(b+5),colname(b),colname(b+3)))  # 1年以上列＝合計との差額
# 当年累計（「（当年）」の見出しのブロック）を、17期の入力済みの月まで
cum=next(ci for ci in range(0,200) if '（当年）' in m.getCellByPosition(ci,1).getString())
filled=[ym for ym in sorted(blocks) if ym>=202511 and (ym<=max(months) or ym in have)]
filled=[ym for ym in filled if ym<=max(months)]
for k in range(5):
    for row in range(0,40):
        cell=m.getCellByPosition(cum+k,row); f=cell.getFormula()
        mm=re.fullmatch(r'=SUM\(((?:[A-Z]+\d+[;,])+[A-Z]+\d+)\)',f)
        if mm:
            sep=';' if ';' in f else ','
            rn=re.match(r'[A-Z]+(\d+)',mm.group(1)).group(1)
            cell.setFormula('=SUM('+sep.join(colname(blocks[ym]+k)+rn for ym in filled)+')')
hd=m.getCellByPosition(cum,1); hd.setString(re.sub(r'～\d{6}','～%d'%max(months),hd.getString()))
doc.calculateAll()
print('当年累計',hd.getString(),'売上',m.getCellByPosition(cum+5,9).getValue(),'粗利',m.getCellByPosition(cum+5,24).getValue())
doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(dst)),(P('FilterName','Calc MS Excel 2007 XML'),))
doc.close(True)
