"""前月の月次報告xlsx → 当月版を作る（LibreOffice UNO）"""
import uno, re, json, sys, os
from com.sun.star.beans import PropertyValue
from com.sun.star.table import CellRangeAddress, CellAddress

def P(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p
def connect():
    ctx=uno.getComponentContext()
    r=ctx.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',ctx)
    c=r.resolve('uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext')
    return c.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop',c)
def col2n(s):
    n=0
    for ch in s: n=n*26+ord(ch)-64
    return n-1
def n2col(n):
    s=''; n+=1
    while n: n,r=divmod(n-1,26); s=chr(65+r)+s
    return s
def used_rows(sh):
    cur=sh.createCursor(); cur.gotoEndOfUsedArea(False); return cur.getRangeAddress().EndRow+1
def find_row(sh,label,cols=(0,),start=0,end=None):
    end=end or used_rows(sh)
    data=sh.getCellRangeByPosition(0,start,max(cols),end-1).getDataArray()
    for i,row in enumerate(data):
        for c in cols:
            if str(row[c]).strip()==label: return start+i
    return None
def cv(v):
    if v is None: return ''
    if isinstance(v,bool): return str(v)
    if isinstance(v,(int,float)): return float(v)
    return str(v)
def arr(block):
    w=max(len(r) for r in block)
    return tuple(tuple(cv(r[i]) if i<len(r) else '' for i in range(w)) for r in block)
def leaf_label(row):  # freee貼付の A列ロジック相当
    for v in row:
        if v not in ('',None): return str(v)
    return ''

def roll(src, dst, m):
    """m: dict of materials"""
    d=connect()
    doc=d.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(src)),'_blank',0,(P('Hidden',True),))
    S=doc.Sheets
    log=[]
    n=m['month_index']; mon=m['month_label']; prevmon=m['prev_month_label']
    # 1. freee貼付(BS)
    for name,c0,block,fcols in [('freee貼付(BS)',2,m['freee_bs'],[(0,1)]),('freee貼付(PL)当年',3,m['freee_pl'],[(0,1),(22,33)])]:
        sh=S.getByName(name); old=used_rows(sh)
        width=max(len(r) for r in block)
        sh.getCellRangeByPosition(c0,0,c0+width-1,max(old,len(block))+5).clearContents(1|2|4|16)  # values,dates,strings,formulas
        rng=sh.getCellRangeByPosition(c0,0,c0+width-1,len(block)-1)
        rng.setDataArray(arr(block))
        if len(block)>old:
            for a,b in fcols:
                srcr=sh.getCellRangeByPosition(a,old-1,b,old-1).getRangeAddress()
                for r in range(old,len(block)):
                    sh.copyRange(sh.getCellByPosition(a,r).getCellAddress(),srcr)
            log.append(f'{name}: 行数 {old}→{len(block)}、数式列を延長')
    # 2. 月数
    S.getByName('月次推移').getCellRangeByName('E1').setValue(n)
    # 3. 新規勘定科目の行追加（BS）
    order=[]; 
    for r in m['freee_bs'][2:]:
        lab=leaf_label(r[0:6]); vals=[v for v in r[6:] if isinstance(v,(int,float))]
        order.append((lab,any(v!=0 for v in vals)))
    targets=[('組替(BS)',0,0),('BS増減分析',0,0),('月次推移',0,find_row(S.getByName('月次推移'),'【BS】') or 0)]
    for sname,lc,start in targets:
        sh=S.getByName(sname)
        existing=set(str(x[0]).strip() for x in sh.getCellRangeByPosition(lc,start,lc,used_rows(sh)).getDataArray())
        for i,(lab,nz) in enumerate(order):
            if not lab or lab in existing or not nz: continue
            # 直前の既存科目を探す
            pred=None
            for j in range(i-1,-1,-1):
                if order[j][0] in existing: pred=order[j][0]; break
            pr=find_row(sh,pred,(lc,),start)
            # 直前科目の上に挿入→直前科目行の数式を複写→ラベルを入替（SUM範囲の内側に入れるため）
            sh.Rows.insertByIndex(pr,1)
            sh.copyRange(sh.getCellByPosition(0,pr).getCellAddress(), sh.getCellRangeByPosition(0,pr+1,200,pr+1).getRangeAddress())
            sh.getCellByPosition(lc,pr).setString(pred); sh.getCellByPosition(lc,pr+1).setString(lab)
            # 値セル（手入力）は新科目側ではクリア
            for c in range(1,200):
                cell=sh.getCellByPosition(c,pr+1)
                if cell.getType().value=='VALUE' and sname!='BS増減分析': cell.setValue(0)
            existing.add(lab); log.append(f'{sname}: 新科目「{lab}」を「{pred}」の下に追加')
    # 4. 組替の手入力欄
    pl=S.getByName('組替(PL)当年'); bs=S.getByName('組替(BS)')
    plc=col2n('T')+n-1; bsc=col2n('P')+n
    def setv(sh,label,col,val,cols=(0,1,2,3,4,5)):
        r=find_row(sh,label,cols)
        if r is None: log.append(f'!! {sh.Name}に「{label}」行なし'); return None
        c=sh.getCellByPosition(col,r)
        if isinstance(val,str): c.setFormula(val)
        else: c.setValue(val)
        return r
    setv(pl,'外注費(派遣費用)',plc,m['haken']); setv(pl,'外注費',plc,-m['haken'])
    setv(pl,'MA会社費用',plc,m['ma'])
    if m.get('fudosan_formula'): setv(pl,'不動産取得税',plc,m['fudosan_formula'])
    rt=find_row(pl,'法人税等',(0,1,2,3,4,5))
    cum=f"SUM('組替(PL)当年'.{n2col(col2n('T'))}{rt+1}:{n2col(plc)}{rt+1})"
    setv(bs,'未払法人税等',bsc,f'={cum}'); setv(bs,'当期純損益金額',bsc,f'=-{cum}')
    rk=find_row(bs,'新株予約権',(0,))
    bs.getCellByPosition(bsc,rk).setValue(bs.getCellByPosition(bsc-1,rk).getValue())
    # 5. 買取再販まとめ
    mt=S.getByName('買取再販まとめ'); cc=3+3*(n-1)+1
    mi=m['matome']
    for row,key in [(5,'人員'),(10,'同日件数'),(11,'通常件数'),(16,'通常売上K'),(21,'通常粗利K'),(51,'滞留件数'),(56,'滞留売上K'),(61,'滞留粗利K')]:
        if mi.get(key) is not None: mt.getCellByPosition(cc,row-1).setValue(mi[key])
    # BX..CA の当月ポインタを3列右へ
    bx,ca=col2n('BX'),col2n('CA')
    def shift(f):
        def rep(mo):
            c=mo.group(2); 
            if bx<=col2n(c)<=ca: return mo.group(0)
            return mo.group(1)+n2col(col2n(c)+3)+mo.group(3)
        f=re.sub(r'(\$?)([A-Z]{1,2})(\$?\d+)',rep,f)
        f=re.sub(r'\*%d\)'%(n-1),'*%d)'%n,f)
        return f
    for r in range(0,70):
        for c in range(bx,ca+1):
            cell=mt.getCellByPosition(c,r); f=cell.getFormula()
            if f.startswith('='): cell.setFormula(shift(f))
    fk=m['fudosan_k']; 
    for addr,v in [('BX21',fk['month']),('BZ21',fk['cum'])]:
        c=mt.getCellRangeByName(addr); base=re.match(r'=(\$?[A-Z]+\$?\d+)',c.getFormula()).group(1)
        c.setFormula(f'={base}+{v}')
    # 6. ラベル
    mo=S.getByName('月次推移')
    mo.getCellRangeByName('V2').setString(f'{prevmon}⇒{mon}増減')
    for r in range(0,used_rows(mo)):
        c=mo.getCellByPosition(col2n('W'),r); s=c.getString()
        if f'【{prevmon}単月】' in s: c.setString(s.replace(f'【{prevmon}単月】',f'【{mon}単月】'))
    vc=col2n('V')
    for r in range(0,used_rows(mo)):
        c=mo.getCellByPosition(vc,r); f=c.getFormula()
        if f.startswith('='):
            c.setFormula(re.sub(r'(?<![A-Z$])(\$?)([G-U])(\$?\d+)',lambda x:x.group(1)+chr(ord(x.group(2))+1)+x.group(3),f))
    S.getByName('BS増減分析').getCellRangeByName('C3').setString(m['bs_title'])
    # 7. 売上明細シート
    sname=f'売上明細({mon})'
    if not S.hasByName(sname):
        idx=[S.getByIndex(i).Name for i in range(S.Count)].index(f'売上明細({prevmon})')
        S.insertNewByName(sname,idx)
    sm=S.getByName(sname)
    pv=m['sales_pivot']; sm.getCellRangeByPosition(0,2,len(pv[0])-1,2+len(pv)-1).setDataArray(arr(pv))
    dt=m['sales_detail']; sm.getCellRangeByPosition(0,99,len(dt[0])-1,99+len(dt)-1).setDataArray(arr(dt))
    doc.calculateAll()
    doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(dst)),(P('FilterName','Calc MS Excel 2007 XML'),))
    doc.close(True)
    return log
