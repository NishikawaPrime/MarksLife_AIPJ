"""10月棚卸の仕上げ：消費税（O・P列）、確認事項、注記"""
import openpyxl, pandas as pd
from openpyxl.styles import PatternFill, Font
res=pd.read_pickle('tax_by_tag_oct.pkl')
p='out/2026.10棚卸_AI作成.xlsx'
wb=openpyxl.load_workbook(p); sp=wb['売上PIVOT']
Y=PatternFill('solid',fgColor='FFF176'); N=PatternFill(fill_type=None)
H=next(r for r in range(1,sp.max_row+1) if sp.cell(r,1).value=='メモタグ')
r=H+1; n=0
while sp.cell(r,1).value:
    t=sp.cell(r,1).value
    sp.cell(r,15,round(-res['仮受消費税'].get(t,0))); sp.cell(r,16,round(res['仮払消費税'].get(t,0)))
    for c in (15,16): sp.cell(r,c).number_format='#,##0'; sp.cell(r,c).fill=(Y if t=='情報2025063776②' else N)
    if t in ('案件0004742','案件0005346'):
        for c in range(1,9): sp.cell(r,c).fill=Y
    r+=1; n+=1
for c in (15,16): sp.cell(H,c).fill=PatternFill('solid',fgColor='DDEBF7')
sp['J1']='黄色の列：AIでは埋められない列（H：手入力調整／U：コメント）。O・P列（消費税）は仕訳帳から按分して計算（黄色の行は要確認）'
sp['J1'].font=Font(bold=True); sp['J1'].fill=Y
ws=wb['AI要確認']; rr=ws.max_row+2
for t in ['【10月の追加の確認事項】',
 '情報2025063776（名古屋市西区江向町 土地）：10月に2分割（②を売却）。先方の10月棚卸の「9月在庫」が分割後（33,474,054円ずつ）になっていたため、それを前月在庫として使用',
 '案件0004742（四條畷市田原台）：先方は売上明細の前月在庫に+12,408,839円を手入力で加算（freee粗利4,571,449円）。AI版は加算せず（freee粗利16,980,288円）',
 '案件0005346（高崎市飯塚町 土地）：freeeで10月30日に売上計上、kintone（買取再販03487）は売却決済実績日が空欄・進捗「売却契約済」のまま。在庫から除外（先方と同じ）',
 '【消費税（売上PIVOTのO・P列）】',
 f'仕訳帳（2025年5月〜2026年10月）から、消費税の行を同じ仕訳の本体行のメモタグで按分して全{n}件に入力（8月の実物で25件中23件一致を確認済みの方法）',
 '情報2025063776②は、仕入の消費税が分割前の「情報2025063776」に付いているため、②の仮払消費税は0円（分割前の案件側に545,262円）',
 '先方の10月棚卸のO・P列は、入力のある25件すべてが8月の値の組と同じ（9月と同様、前の値が残っている可能性）']:
    ws.cell(rr,1,t).fill=Y; rr+=1
wb.save(p); print('件数',n,'見出し行',H)
