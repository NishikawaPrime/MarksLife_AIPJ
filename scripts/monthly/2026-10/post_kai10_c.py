"""10月棚卸：C列（手入力調整）の扱い＝空欄のまま＋支払い漏れ候補の提示"""
import openpyxl, json
from openpyxl.styles import PatternFill, Font, Alignment
pv=json.load(open('/tmp/claude-0/-home-claude/826d9c02-0d73-542f-934f-e59e83a2b43f/scratchpad/pv.json'))
cand=sorted([x for x in pv if x[6]>=1_000_000],key=lambda x:-x[6])
p='out/2026.10棚卸_AI作成.xlsx'
wb=openpyxl.load_workbook(p); sp=wb['売上PIVOT']
Y=PatternFill('solid',fgColor='FFF176')
H=next(r for r in range(1,sp.max_row+1) if sp.cell(r,1).value=='メモタグ')
rowof={sp.cell(r,1).value:r for r in range(H+1,sp.max_row+1) if sp.cell(r,1).value}
for x in cand: sp.cell(rowof[x[0]],8).fill=Y
ws=wb['AI要確認']; rr=ws.max_row+2
ws.cell(rr,1,'【売上PIVOT H列「C：支払い漏れ／前期以前登記費用・火災保険などの手入力」】').font=Font(bold=True)
ws.cell(rr+1,1,'AIでは金額を決められないため空欄のまま。必要な案件があれば会社で記入してください（8〜10月の月次報告では入力なし、11〜4月は月1〜3件）')
ws.cell(rr+2,1,'支払い漏れの候補：freee粗利がkintone粗利より100万円以上多い案件（freee側に費用が入っていない可能性）。H列を黄色にしています')
ws.cell(rr+3,1,'内容'); ws.cell(rr+3,2,'メモタグ'); ws.cell(rr+3,3,'金額')
for i,x in enumerate(cand):
    note=f"{x[1]}：freee粗利 {x[3]:,.0f}円、kintone粗利 {x[5]:,.0f}円、差 {x[6]:,.0f}円"
    if x[0]=='案件0004742': note+='（先方は前月在庫に+12,408,839円を手入力で加算して調整済み）'
    ws.cell(rr+4+i,1,note); ws.cell(rr+4+i,2,x[0]); c=ws.cell(rr+4+i,3,round(x[6])); c.number_format='#,##0'
wb.save(p); print(len(cand),[x[0] for x in cand])
