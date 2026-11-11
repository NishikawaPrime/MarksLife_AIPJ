import uno, os
from com.sun.star.beans import PropertyValue
def P(n,v): p=PropertyValue(); p.Name=n; p.Value=v; return p
ctx=uno.getComponentContext(); r=ctx.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',ctx)
c=r.resolve('uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext')
desk=c.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop',c)
doc=desk.loadComponentFromURL('file:///home/claude/oct/run/uri_out.xlsx','_blank',0,(P('Hidden',True),))
w=doc.Sheets.getByName('AI作業メモ')
cur=w.createCursor(); cur.gotoEndOfUsedArea(False); r0=cur.RangeAddress.EndRow+2
lines=[['売上構成分析 10月分の追加（AI作成、2026/11/11）'],
 ['17期_買取再販明細・買取再販明細','2026年10月180件を追加（AI作成の10月棚卸の売上明細。件数・売上は先方の10月棚卸と一致）'],
 ['不動産取得税','10月分85件（合計11,853,800円、月次報告と一致）'],
 ['買取再販粗利','10月の同日／通常の売上・粗利は月次報告（AI作成10月版）の買取再販まとめから'],
 ['まとめ','当年累計を202511～202610に延長。10月の粗利の行は7月と同じ形にそろえた'],
 ['確認してほしい点','案件0005346（高崎市飯塚町 土地）はkintoneに売却決済日がないため、freeeの売上計上日10/30から112日として集計（先方の10月棚卸と同じ）'],
 ['','案件0004742（四條畷市田原台）は、先方の10月棚卸で前月在庫に+12,408,839円が手入力で加算されているが、根拠がないためAI版では加算していない。このファイルの10月の粗利は、その分だけ先方より大きい'],
 ['','10月まで含めると全体の粗利率は26.7%→24.1%（▲2.6pt）。同日決済の構成比45.7%→35.0%、同日決済の粗利率30.8%→28.1%、3ヶ月未満の粗利率31.8%→24.4%。まとめシートのコメント欄（7月時点）は書き換えていない']]
for i,row in enumerate(lines):
    for j,v in enumerate(row): w.getCellByPosition(j,r0+i).setString(v)
w.getCellByPosition(0,r0).CharWeight=150
doc.storeToURL('file:///home/claude/oct/run/uri_final.xlsx',(P('FilterName','Calc MS Excel 2007 XML'),))
doc.close(True); print('ok')
