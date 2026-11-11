import sys,datetime,openpyxl
sys.path.insert(0,'.')
from tanaoroshi_builder import zen
from openpyxl.styles import PatternFill,Font
Y=PatternFill('solid',fgColor='FFF176')
def highlight(path,mon,res):
    wb=openpyxl.load_workbook(path)
    sp=wb['売上PIVOT']
    H=next(r for r in range(1,sp.max_row+1) if sp.cell(r,1).value=='メモタグ')
    r=H+1
    while sp[f'A{r}'].value:
        for c in 'HOPU': sp[f'{c}{r}'].fill=Y
        r+=1
    for c in 'HOPU': sp[f'{c}{H}'].fill=Y
    sp['J1']='黄色の列：AIでは埋められない列（H：手入力調整／O・P：消費税はカスタムレポート「BSメモタグ集計」が必要／U：コメント）'
    sp['J1'].font=Font(bold=True); sp['J1'].fill=Y
    flagged={a[1] for a in res['alerts']}
    inv=wb[f'{zen(mon)}月在庫']; n=0
    for row in inv.iter_rows(min_row=3):
        if row[0].value in flagged:
            for c in row[:6]: c.fill=Y
            n+=1
    inv['G1']='黄色の行：kintoneで「没」など、在庫に残すか外すかの判断が必要な案件（AI要確認シート参照）'; inv['G1'].font=Font(bold=True); inv['G1'].fill=Y
    wb.save(path); return n
