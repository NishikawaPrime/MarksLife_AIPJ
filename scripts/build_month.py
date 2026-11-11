"""月次報告の作成。mode='roll'：前月版→当月版 / mode='refresh'：同じ月のファイルを最新データで作り直す"""
import uno, re, os, sys
sys.path.insert(0,'.')
from roll_report import connect,P,col2n,n2col,used_rows,find_row,arr,leaf_label
MONTHS=['11月','12月','1月','2月','3月','4月','5月','6月','7月','8月','9月','10月']
def ref(r,c): return f"$'手入力値'.{n2col(c)}{r+1}"

def write_manual(S,n,mv,log):
    if S.hasByName('手入力値'): S.removeByName('手入力値')
    idx=[S.getByIndex(i).Name for i in range(S.Count)].index('組替(BS)')
    S.insertNewByName('手入力値',idx); sh=S.getByName('手入力値')
    hdr=['項目','反映先','出典','期首']+MONTHS+['メモ']
    rows=[['手入力値（freee本体の月次推移以外から反映する数値）']+['']*16,
          ['※組替シート・買取再販まとめはこのシートを参照しています。数値を直すときはこのシートだけを直してください。']+['']*16,
          hdr]
    def line(name,dest,src,vals,memo='',kishu=''):
        return [name,dest,src,kishu]+[('' if v is None else v) for v in vals]+[memo]
    rows.append(line('外注費(派遣費用)','組替(PL)当年：外注費(派遣費用)に＋、外注費に－','本体の仕訳帳：外注費のうち品目「【外注費】人材派遣」',mv['haken']))
    rows.append(line('不動産取得税','組替(PL)当年：不動産取得税（売上原価へ振替）','本体の仕訳帳：仕入高（租税公課）の品目「不動産取得税（確定）」＋「不動産取得税（概算）」',mv['fudosan']))
    rows.append(line('MA会社費用','組替(PL)当年：MA会社費用','子会社マークスライフＭ＆Ａ株式会社の月次推移（PL）の「販売管理費 計」',mv['ma'],'雑損失（営業外）は含めない'))
    accr=['' if (i>=n) else 0 for i in range(12)]
    rows.append(line('MA会社費用の未払費用計上','組替(BS)：未払費用に＋、当期純損益金額に－','MA会社費用の累計（上の行）',accr,'9月の取締役会資料に合わせ、9月末からMA会社費用の累計を未払費用に計上'))
    rows.append(line('新株予約権','組替(BS)：新株予約権','前月版からの引継ぎ（freeeに計上なし）',[24000 if i<n else None for i in range(12)],'',24000))
    rows.append(line('営業人員','買取再販まとめ：人員（当年）','17期_人員の「まとめ」シート：身分「社員」×機能「営業」の当年',mv['jinin']))
    sh.getCellRangeByPosition(0,0,16,len(rows)-1).setDataArray(arr(rows))
    R={'haken':3,'fudosan':4,'ma':5,'accr':6,'kabu':7,'jinin':8}
    # 未払費用計上：計上開始月以降は MA の累計
    m0=mv['ma_accr_month']
    for i in range(m0-1,n):
        sh.getCellByPosition(4+i,R['accr']).setFormula(f'=SUM({n2col(4)}{R["ma"]+1}:{n2col(4+i)}{R["ma"]+1})')
    for c,w in [(0,5200),(1,9000),(2,14000),(16,12000)]: sh.getColumns().getByIndex(c).Width=w
    for c in range(3,16): sh.getColumns().getByIndex(c).Width=2600
    sh.getCellRangeByPosition(3,3,15,len(rows)-1).NumberFormat=doc_fmt(S)
    return R

_fmt=None
def doc_fmt(S):
    return _fmt

def link_manual(S,n,R,log):
    pl=S.getByName('組替(PL)当年'); bs=S.getByName('組替(BS)'); mt=S.getByName('買取再販まとめ')
    T=col2n('T')
    def plrow(label):
        r=find_row(pl,label,(0,1,2,3,4,5)); assert r is not None,label; return r
    rh,rg,rf,rm=plrow('外注費(派遣費用)'),plrow('外注費'),plrow('不動産取得税'),plrow('MA会社費用')
    for i in range(12):
        c=T+i
        pl.getCellByPosition(c,rh).setFormula(f'={ref(R["haken"],4+i)}')
        pl.getCellByPosition(c,rg).setFormula(f'=-{ref(R["haken"],4+i)}')
        pl.getCellByPosition(c,rf).setFormula(f'={ref(R["fudosan"],4+i)}')
        pl.getCellByPosition(c,rm).setFormula(f'={ref(R["ma"],4+i)}')
    rt=find_row(pl,'法人税等',(0,1,2,3,4,5))
    Pb=col2n('P')
    rmi=find_row(bs,'未払費用',(0,)); rnet=find_row(bs,'当期純損益金額',(0,)); rtax=find_row(bs,'未払法人税等',(0,)); rk=find_row(bs,'新株予約権',(0,))
    bs.getCellByPosition(Pb,rk).setFormula(f'={ref(R["kabu"],3)}')
    for i in range(12):
        c=Pb+1+i
        cum=f"SUM($'組替(PL)当年'.{n2col(T)}{rt+1}:{n2col(T+i)}{rt+1})"
        bs.getCellByPosition(c,rk).setFormula(f'={ref(R["kabu"],4+i)}')
        bs.getCellByPosition(c,rmi).setFormula(f'={ref(R["accr"],4+i)}')
        if i<n:
            bs.getCellByPosition(c,rtax).setFormula(f'={cum}')
            bs.getCellByPosition(c,rnet).setFormula(f'=-{cum}-{ref(R["accr"],4+i)}')
        else:
            bs.getCellByPosition(c,rtax).setString(''); bs.getCellByPosition(c,rnet).setString('')
    for i in range(12):
        if i<n: mt.getCellByPosition(3+3*i+1,4).setFormula(f'={ref(R["jinin"],4+i)}')
    log.append('組替(PL)当年・組替(BS)・買取再販まとめの手入力欄（派遣費用・不動産取得税・MA会社費用・新株予約権・未払法人税等・営業人員）を「手入力値」シートの参照に置き換え')

def build(src,dst,m,mode):
    d=connect()
    doc=d.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(src)),'_blank',0,(P('Hidden',True),))
    S=doc.Sheets; log=[]
    global _fmt
    _fmt=doc.NumberFormats.queryKey('#,##0',doc.CharLocale,False)
    if _fmt==-1: _fmt=doc.NumberFormats.addNew('#,##0',doc.CharLocale)
    n=m['month_index']; mon=m['month_label']; prevmon=m['prev_month_label']
    # 1. freee貼付
    for name,c0,block,fcols in [('freee貼付(BS)',2,m['freee_bs'],[(0,1)]),('freee貼付(PL)当年',3,m['freee_pl'],[(0,1),(22,33)])]:
        sh=S.getByName(name); old=used_rows(sh)
        width=max(len(r) for r in block)
        sh.getCellRangeByPosition(c0,0,c0+width-1,max(old,len(block))+5).clearContents(1|2|4|16)
        sh.getCellRangeByPosition(c0,0,c0+width-1,len(block)-1).setDataArray(arr(block))
        for a,b in fcols:
            # 検索キー列の数式が貼付ブロックの最終行まで届いているか確認し、足りなければ延長する
            col=sh.getCellRangeByPosition(a,0,a,max(old,len(block))).getFormulaArray()
            lastf=max((i for i,row in enumerate(col) if str(row[0]).startswith('=')),default=None)
            if lastf is not None and lastf<len(block)-1:
                srcr=sh.getCellRangeByPosition(a,lastf,b,lastf).getRangeAddress()
                for r in range(lastf+1,len(block)): sh.copyRange(sh.getCellByPosition(a,r).getCellAddress(),srcr)
    log.append(f'freee貼付(BS)・freee貼付(PL)当年に最新のfreee月次推移を貼付（{mon}まで）')
    S.getByName('月次推移').getCellRangeByName('E1').setValue(n)
    # 2. 新規BS科目
    order=[]
    for r in m['freee_bs'][2:]:
        lab=leaf_label(r[0:6]); vals=[v for v in r[6:] if isinstance(v,(int,float))]
        order.append((lab,any(v!=0 for v in vals)))
    added=[]
    for sname,lc,start in [('組替(BS)',0,0),('BS増減分析',0,0),('月次推移',0,find_row(S.getByName('月次推移'),'【BS】') or 0)]:
        sh=S.getByName(sname)
        existing=set(str(x[0]).strip() for x in sh.getCellRangeByPosition(lc,start,lc,used_rows(sh)).getDataArray())
        for i,(lab,nz) in enumerate(order):
            if not lab or lab in existing or not nz: continue
            pred=next((order[j][0] for j in range(i-1,-1,-1) if order[j][0] in existing),None)
            pr=find_row(sh,pred,(lc,),start)
            sh.Rows.insertByIndex(pr,1)
            sh.copyRange(sh.getCellByPosition(0,pr).getCellAddress(), sh.getCellRangeByPosition(0,pr+1,200,pr+1).getRangeAddress())
            sh.getCellByPosition(lc,pr).setString(pred); sh.getCellByPosition(lc,pr+1).setString(lab)
            for c in range(1,200):
                cell=sh.getCellByPosition(c,pr+1)
                if cell.getType().value=='VALUE' and sname!='BS増減分析': cell.setValue(0)
            existing.add(lab)
            if lab not in added: added.append(lab)
    if added: log.append('freeeに新しく出てきたBS科目を追加：'+'、'.join(added))
    # 3. 手入力値シート
    R=write_manual(S,n,m['manual'],log); link_manual(S,n,R,log)
    # 4. 買取再販まとめ
    mt=S.getByName('買取再販まとめ'); cc=3+3*(n-1)+1; mi=m['matome']
    for row,key in [(10,'同日件数'),(11,'通常件数'),(16,'通常売上K'),(21,'通常粗利K'),(51,'滞留件数'),(56,'滞留売上K'),(61,'滞留粗利K')]:
        if mi.get(key) is not None: mt.getCellByPosition(cc,row-1).setValue(mi[key])
    if mode=='roll':
        bx,ca=col2n('BX'),col2n('CA')
        def shift(f):
            def rep(mo):
                c=mo.group(2)
                if bx<=col2n(c)<=ca: return mo.group(0)
                return mo.group(1)+n2col(col2n(c)+3)+mo.group(3)
            f=re.sub(r'(\$?)([A-Z]{1,2})(\$?\d+)',rep,f)
            return re.sub(r'\*%d\)'%(n-1),'*%d)'%n,f)
        for r in range(0,70):
            for c in range(bx,ca+1):
                cell=mt.getCellByPosition(c,r); f=cell.getFormula()
                if f.startswith('='): cell.setFormula(shift(f))
    fk=m['fudosan_k']
    for addr,v in [('BX21',fk['month']),('BZ21',fk['cum'])]:
        c=mt.getCellRangeByName(addr); base=re.match(r'=(\$?[A-Z]+\$?\d+)',c.getFormula()).group(1)
        c.setFormula(f'={base}+{v}')
    # 5. ラベル（月を進めるときだけ）
    if mode=='roll':
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
    # 6. 売上明細
    sname=f'売上明細({mon})'
    if not S.hasByName(sname):
        idx=[S.getByIndex(i).Name for i in range(S.Count)].index(f'売上明細({prevmon})')
        S.insertNewByName(sname,idx)
    sm=S.getByName(sname)
    sm.getCellRangeByPosition(0,2,40,used_rows(sm)+5).clearContents(1|2|4|16)
    pv=m['sales_pivot']; sm.getCellRangeByPosition(0,2,len(pv[0])-1,2+len(pv)-1).setDataArray(arr(pv))
    dt=m['sales_detail']; sm.getCellRangeByPosition(0,99,len(dt[0])-1,99+len(dt)-1).setDataArray(arr(dt))
    log.append(f'{sname}シートを棚卸の生データから案件別に作成')
    doc.calculateAll()
    doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(dst)),(P('FilterName','Calc MS Excel 2007 XML'),))
    doc.close(True)
    return log
