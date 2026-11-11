import sys,json; sys.path.insert(0,'.')
from board_deck import build
from board_bs import bs_params
from cf_loan import cf_lines, inv_gp, sheet_bond
T='/mnt/user-data/uploads/Downloads/20261016取締役会資料.pptx'
LOAN='/home/claude/oct/in3/x1b.xlsx'; KIN='/home/claude/oct/in2/daicho.csv'
b=json.load(open('/home/claude/oct/in2/budget10.json')); budget=dict(m=b['m'],c=b['c'])
bs=bs_params('n10_5.xlsx',12,'10月',309)
loan,cash,_=cf_lines(LOAN,'n10_5.xlsx','10月',True,sheet_bond(LOAN))
gp,cnt=inv_gp(KIN,'2026-10-31')
bs.update(loan_lines=loan,cash_lines=cash,kintone_lines=[f'（在庫粗利：{gp:,}）','Kintone成約予定ベース'],loan_red=(),cash_red=(),kintone_red=())
build(T,'d10.pptx','n10_5.xlsx',12,'10',budget,bs)
print('ok',gp,cnt)
# 表の数値が改行されないよう、はみ出すセルだけフォントを縮める（毎月必ず実行）
from pptx import Presentation
import deck_fit
_p=Presentation('d10.pptx'); print('縮めたセル',deck_fit.fit(_p)); _p.save('d10.pptx')
