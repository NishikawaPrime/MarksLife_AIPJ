import uno,os,sys,csv
sys.path.insert(0,'/home/claude/oct/run')
from roll_report import connect,P,col2n
import pandas as pd, openpyxl
src,dst=sys.argv[1:3]
K=lambda x:f"{x/1000:+,.0f}K".replace('+-','△').replace('-','△')
M=lambda x:f"{x/1e6:+,.0f}M".replace('-','△')
OKU=lambda x:(f"＋{x/1e8:.1f}億" if x>=0 else f"△{-x/1e8:.1f}億")
rows=list(csv.reader(open('/home/claude/oct/in/f3.csv',encoding='cp932'))); h=rows[1]
i0,i10=h.index('期首'),h.index('2026-10')
def val(name,i):
    for r in rows[2:]:
        if next((x for x in r[:6] if x),'')==name: return int(r[i] or 0)
    return 0
def tot(key,i): return sum(int(r[i] or 0) for r in rows[2:] if any(key in x for x in r[:6]))
pbt=283671150
cash=1512123484-1134567952
inv=(val('販売用不動産',i10)+val('仕掛販売用不動産',i10))-(val('販売用不動産',i0)+val('仕掛販売用不動産',i0))
bon=val('賞与引当金',i10)-val('賞与引当金',i0)
loan=4761177577-2086070165
bond=(val('社債',i10)+val('１年以内返済予定社債',i10))-(val('社債',i0)+val('１年以内返済予定社債',i0))
dep=val('敷金',i10)-val('敷金',i0)
fa=(val('有形固定資産 計',i10)+val('無形固定資産 計',i10))-(val('有形固定資産 計',i0)+val('無形固定資産 計',i0))
inv_cf=-dep-fa; fin=loan+bond; op=cash-pbt-inv_cf-fin; oth=op+inv-bon
tsu0,tsu1=tot('通知預金',i0),tot('通知預金',i10); tei0,tei1=tot('定期',i0),tot('定期',i10)
wb=openpyxl.load_workbook('/home/claude/oct/run/out/2026.10棚卸_AI作成.xlsx',read_only=True,data_only=True)
v=[r[2] for r in wb['１０月在庫'].iter_rows(min_row=3,values_only=True) if isinstance(r[2],(int,float))]
big=[x for x in v if x>50_000_000]; n,s=len(v),sum(v); assert s==val('販売用不動産',i10),(s,)
L=pd.read_pickle('/home/claude/oct/c/L.pkl')
yan=L[(L.ym=='2026-10')&(L.acct=='附属設備')&(L.amt>0)&L.item.str.contains('ヤンマー',na=False)].amt.sum()
gifu=L[(L.ym=='2026-10')&(L.acct=='附属設備')&(L.amt>0)&L.item.str.contains('岐阜支店内装',na=False)].amt.sum()
gifue=L[(L.ym=='2026-10')&(L.acct=='附属設備')&(L.amt>0)&L.item.str.contains('岐阜支店電気',na=False)].amt.sum()
GRN,YEL=0xC8E6C9,0xFFF176
U={ # 行:(科目,新テキスト or (旧,新)置換リスト, 追記行, 色)
133:('現預金',f"サマリ：利益{OKU(pbt)}、営業{OKU(op)}、投資{OKU(inv_cf)}、借入{OKU(fin)}\n税前利益：{M(pbt)}\n営業：在庫{M(-inv)}、賞与引当金{M(bon)}、その他（税金・前受金等）{M(oth)}\n投資：敷金{M(-dep)}、新支店他{M(-fa)}\n財務：借入金{M(loan)}、社債{M(bond)}\n\n通知預金：{tsu1/1000:,.0f}K({K(tsu1-tsu0)})\n定期積金：{tei1/1000:,.0f}K({K(tei1-tei0)})",None,GRN),
135:('販売用不動産',f"期首：203件 @10,025K\n10月末：{n}件 ＠{s/n/1000:,.0f}K（50百万超物件{len(big)}件除く：{n-len(big)}件、@{(s-sum(big))/(n-len(big))/1000:,.0f}K）",None,GRN),
154:('リース資産（有形）',[('△5,551K','△6,689K')],None,GRN),
155:('建物',"【取得】10月：リブゼ横浜サザンティアラ201号室（建物部分）+12,374K（取得目的・用途：要記入）\n【減価償却費】 △419K",None,YEL),
156:('附属設備',[('【減価償却費】 △11,544K',
 f"         9月：神奈川大和支店（移転）内装+5,425K・B工事+2,275K・消防設備+866K、高崎支店内装+1,950K、岐阜支店B工事+960K\n         10月：ヤンマービル（本社）内装・設備工事{K(yan)}（受贈益22,015Kの設備譲渡分か：要確認）、岐阜支店内装{K(gifu)}・電気{K(gifue)}\n【減価償却費】 △14,325K"),
 ('【除却】 横浜支店(プラウド馬車道)△2,420K','【除却】 横浜支店(プラウド馬車道)△2,420K\n【減損】 9月：新宿支店（法人営業部）移転△7,161K、神奈川大和支店移転△2,525K')],None,YEL),
157:('構築物',[('△63K','△76K')],None,GRN),
159:('工具器具備品',[('【減価償却費】 △2,588K     【除却損】浜松支店△82K',
 "          9月：DX推進部MacBookPro+1,123K、船橋支店フリーデスク+265K、神奈川大和支店ロッカー+234K\n          10月：ヤンマービル（本社）什器+1,302K、岐阜支店複合機+750K、豊岡出石リフォーム資材（五右衛門風呂）+255K\n【減価償却費】 △3,375K     【除却損】浜松支店△82K     【減損】9月：監視カメラ（神奈川大和・新宿）△161K")],None,GRN),
160:('一括償却資産',[('【減価償却費】 △1,721K',"          9月：神奈川大和支店UTM+168K\n          10月：ヤンマービル（本社）内装附随+326K、豊岡出石1号店テーブル+229K\n【減価償却費】 △2,631K")],None,GRN),
161:('土地',None,"10月：リブゼ横浜サザンティアラ201号室+19,466K（取得目的・用途：要記入）",YEL),
162:('建設仮勘定',None,"9月：豊岡出石1号店 改修工事前金+14,000K（徳網建設）、2号店 改修工事前金+8,000K（エス・ビレッジ）\n10月：豊岡出石2号店+8,000K（エス・ビレッジ）",GRN),
165:('リース資産（無形）',[('△410K','△492K')],None,GRN),
166:('ソフトウェア',[('△1,784K','△2,141K')],None,GRN),
169:('敷金',None,"新宿支店（移転）+9,175K、新潟支店+4,518K、甲府支店+694K",GRN),
220:('賞与引当金',None,"9〜10月：決算引当金平準化+164,225K（10月分 インセン+108,907K、BO賞与+8,155K）、感謝手当戻入△6,674K\n※10月の大幅な積み増しの理由：要記入",YEL),
228:('社債',"あいち銀行 152,000K（9月 一部償還△16,000K）、浜松いわた信用金庫 50,000K\n※このほか1年以内償還予定 32,000K（あいち銀行）",None,GRN),
230:('借入金',"【借入+9,389M、返済△6,708M】\n◇PJ資金+1,940M（借入+7,669M、返済△5,729M）\n◇運転資金 +624M（借入+896M、返済△272M）\n◇設備 +68M（借入+90M、返済△22M）\n◇賞与 +49M（借入+484M、返済△435M）\n◇納税 +±0M（借入+250M、返済△250M）\n※借入金返済明細の10月末残高 4,766,378K と帳簿 4,761,178K に5,200Kの差：要確認",None,YEL),
}
d=connect(); doc=d.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(src)),'_blank',0,(P('Hidden',True),))
sh=doc.Sheets.getByName('BS増減分析'); E=col2n('E')
for r,(lab,new,add,col) in U.items():
    a=sh.getCellByPosition(0,r-1).getString().strip(); assert a==lab,(r,a,lab)
    c=sh.getCellByPosition(E,r-1); t=c.getString()
    if isinstance(new,str): t=new
    elif isinstance(new,list):
        for o,nw in new: assert o in t,(r,o); t=t.replace(o,nw)
    if add: t=t+'\n'+add
    c.setString(t); c.CellBackColor=col
for a in ('E1','E2'):
    c=sh.getCellRangeByName(a)
    if not c.getString():
        c.setString('黄色：会社にしか書けない内容（要記入）を含むコメント／緑：AIが10月末の数字で更新したコメント（要確認）'); c.CellBackColor=YEL; break
doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(dst)),(P('FilterName','Calc MS Excel 2007 XML'),))
doc.close(True); print('ok',op,oth,inv_cf,fin,tsu0,tsu1,tei0,tei1,len(big))
