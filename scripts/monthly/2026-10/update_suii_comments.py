import uno,os,sys,json
sys.path.insert(0,'/home/claude/oct/run')
from roll_report import connect,P,col2n
src,dst=sys.argv[1:3]
d=connect(); doc=d.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(src)),'_blank',0,(P('Hidden',True),))
sh=doc.Sheets.getByName('月次推移'); W=col2n('W'); YEL=0xFFF176
def lab(r):
    for c in range(0,6):
        s=sh.getCellByPosition(c,r-1).getString().strip()
        if s: return s
    return ''
# 行番号: (科目名の確認用, 追記する行のリスト, 置換(旧→新) , 黄色)
U={
 53:('給料手当',[],('5月+9人','5月+9人、6月+4人、7月+2人、8月+13人、9月+4人、10月+5人'),False),
 56:('福利厚生費',['9月：社内イベント（TKP）2,013K'],None,False),
 58:('賞与引当金繰入',['9月：決算引当金平準化+47,162K、感謝手当△3,337K','10月：決算引当金平準化+117,063K（インセンティブ108,907K、BO賞与8,155K）、感謝手当△3,337K','　※10月の大幅増の理由（期末の引当見直し等）：要記入'],None,True),
 63:('外注費',['9月：ikiteru 1,080K、Labbit（マーケティング）1,060K','10月：STRARTS 4,915K（内容：要記入）'],None,True),
 64:('広告宣伝費',['9月：博展（イベント出展）3,634K（イベント名：要記入）','10月：東京博善（イベント出展）2,750K（イベント名：要記入）'],None,True),
 67:('支払報酬料',['9月・10月：三優監査法人 各2,300K'],None,False),
 72:('租税公課',[],('8月：14,208K)','8月：14,208K、9月：21,318K、10月：25,134K)'),False),
 73:('減価償却費',['8月〜：月2,900K〜3,392Kに増加（増加の内容：要記入）'],None,True),
 76:('旅費交通費',['10月：タイムズカー4,901K（9月4,170K）、ETC 2,026K、タイムズパーキング1,068K、ガソリン1,014K'],None,False),
 81:('地代家賃',['9月：新潟支店+711K、神奈川大和支店+366K、神戸支店+245K、浜松支店+130K','10月：甲府支店+350K（新規）、新潟支店△355K、神奈川大和支店△277K'],None,False),
 99:('支払利息',['10月：PJ借入の返済（1,670M）が多く、戻し利息△5,165K（33件。9月は△4,036K）','　※減少の主因：要確認'],None,True),
 101:('融資手数料',['9月：セゾンファンデックス 1,545K','10月：三十三銀行 1,360K、筑波銀行（運転資金）1,300K、セゾンファンデックス 1,061K'],None,False),
 265:('前払費用',['10月：火災保険（2026/11〜2027/09分）+9,809K、STRARTS△5,407K'],None,False),
 268:('仮払金',['9月：新宿支店・新潟支店の敷金を敷金へ振替△13,693K','10月：労働保険 5,405K'],None,False),
 284:('土地',['10月：リブゼ横浜サザンティアラ201号室+19,466K（用途：要記入）'],None,True),
 285:('建設仮勘定',['9月：豊岡出石1号店 改修工事前金14,000K（徳網建設）、2号店 改修工事前金8,000K（エス・ビレッジ）','10月：豊岡出石2号店 8,000K（エス・ビレッジ）'],None,False),
 292:('敷金',['9月：新宿支店（移転）+9,175K、新潟支店+4,518K（仮払金から振替）','10月：甲府支店+694K'],None,False),
 319:('未払金',['10月：マークスライフＭ＆Ａ株式会社の設立資本金10,000Kを支払'],None,False),
 331:('未払費用',['9月：MA会社費用の累計（7〜9月）を未払計上+12,354K（組替）、インセンティブ社会保険料+3,273K、不動産取得税+3,251K','10月：インセンティブ社会保険料+7,042K、MA会社費用+2,641K（組替）'],None,False),
 337:('前受金',['10月：売却決済で手付金を充当し△51,005K'],None,False),
 338:('預り金',['10月：源泉所得税+8,069K（10/30の支払分。内容：要記入）'],None,True),
 343:('賞与引当金',['9月：決算引当金平準化+47,162K、感謝手当△3,337K','10月：決算引当金平準化+117,063K、感謝手当△3,337K','　※10月の大幅増の理由：要記入'],None,True),
 351:('社債',['9月：あいち銀行私募債 一部償還△16,000K'],None,False),
 353:('借入金',['9月：PJ+447M、運転資金+72M、返済△78M','10月：PJ△1,137M、運転資金+83M、返済△43M'],None,False),
 83:('寄付金',['10月：豊岡市への寄付 5,000K（内容：要記入）'],None,True),
 109:('受贈益',['10月：ヤンマービル（本社）設備譲渡 22,015K、ほか31,840K（内容：要記入）'],None,True),
 301:('長期貸付金',['10月：マークスライフＭ＆Ａ株式会社への貸付 30,000K（貸付期間2026/07/10〜2032/10/31、利息は子会社側で139K計上）'],None,False),
}
log=[]
for r,(name,add,rep,yel) in U.items():
    assert lab(r)==name, (r,lab(r),name)
    c=sh.getCellByPosition(W,r-1); t=c.getString()
    if rep: assert rep[0] in t,(r,rep[0]); t=t.replace(rep[0],rep[1])
    if add: t=(t.rstrip('\n')+'\n' if t else '')+'\n'.join(add)
    c.setString(t)
    if yel: c.CellBackColor=YEL
    log.append((r,name,yel))
n=sh.getCellRangeByName('W1'); n.setString('黄色：会社にしか書けない内容（要記入）を含むコメント'); n.CellBackColor=YEL
doc.calculateAll()
doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(dst)),(P('FilterName','Calc MS Excel 2007 XML'),))
doc.close(True)
print(len(log),'件更新、黄色',sum(1 for x in log if x[2]))
