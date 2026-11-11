"""前月の棚卸ファイル＋当月の生データ → 当月の棚卸ファイル（同じシート構成・数式）"""
import openpyxl, datetime, sys, copy
sys.path.insert(0,'.')
from kai_engine import *
from openpyxl.utils import get_column_letter as gl
from openpyxl.styles import PatternFill, Font
FW='０１２３４５６７８９'
def zen(n): return ''.join(FW[int(d)] for d in str(n))
def clear(ws, keep_rows=0):
    for row in ws.iter_rows(min_row=keep_rows+1):
        for c in row: c.value=None
def copy_raw(src_ws, dst_ws, c1=1, c2=None):
    clear(dst_ws)
    for row in src_ws.iter_rows(min_col=c1,max_col=c2 or src_ws.max_column):
        for c in row:
            if c.value is not None: dst_ws.cell(c.row,c.column,c.value)
from openpyxl.styles import Border, Side, Alignment
def style_pivot(sp,pv_last,tot_row,FR,H,d1,d2):
    """先方ファイル（ピボットテーブルの見た目）に合わせた書式"""
    # テンプレートの書式の残り（罫線・塗り・フォント）を消してからそろえる
    NOF=PatternFill(fill_type=None); NB=Border()
    last=max(d2,sp.max_row)
    for r in range(3,last+1):
        for c in range(1,27):
            cl=sp.cell(r,c); cl.border=NB; cl.font=Font(name='游ゴシック',size=11)
            if r<H: cl.fill=NOF
    HDR=PatternFill('solid',fgColor='DDEBF7'); TOT=PatternFill('solid',fgColor='DDEBF7'); SUM=PatternFill('solid',fgColor='E2EFDA'); KEY=PatternFill('solid',fgColor='FFF2CC')
    thin=Side(style='thin',color='9BC2E6'); dbl=Side(style='double',color='5B9BD5')
    for c in range(1,5):
        sp.cell(3,c).font=Font(bold=True)
        h=sp.cell(4,c); h.fill=HDR; h.font=Font(bold=True); h.border=Border(bottom=thin)
        t=sp.cell(tot_row,c); t.fill=TOT; t.font=Font(bold=True); t.border=Border(top=dbl)
    for r in range(5,tot_row):
        for c in range(2,5): sp.cell(r,c).number_format='#,##0'
        sp.cell(r,1).border=Border(bottom=Side(style='hair',color='D9D9D9'))
    for c in range(2,5): sp.cell(tot_row,c).number_format='#,##0'
    # 集計ブロック
    for c in range(1,9):
        sp.cell(FR,c).font=Font(bold=True)
    for r in (FR+1,FR+2):
        sp.cell(r,4).fill=PatternFill('solid',fgColor='DDEBF7'); sp.cell(r,4).font=Font(bold=True)
        sp.cell(r,7).number_format='0.0%'
        for c in (8,): sp.cell(r,c).number_format='#,##0_);[Red](#,##0)'
    for c in range(3,9): sp.cell(FR+4,c).fill=SUM; sp.cell(FR+4,c).font=Font(bold=True)
    sp.cell(FR+4,7).number_format='0.0%'
    for c in range(3,7): sp.cell(FR+4,c).number_format='#,##0_);[Red](#,##0)'
    for c in range(2,5): sp.cell(FR,c).number_format='#,##0_);[Red](#,##0)'
    # 売上明細の見出し
    for c in range(1,25):
        h=sp.cell(H,c); h.fill=HDR; h.font=Font(bold=True); h.alignment=Alignment(wrap_text=True,vertical='center',horizontal='center'); h.border=Border(bottom=thin)
    sp.row_dimensions[H].height=52.8
    for r in range(d1,d2+1): sp.cell(r,6).fill=KEY
def build(template, out, raw_wb, prev_inv_sheet, mon, prevmon, base_date):
    wb=openpyxl.load_workbook(template)
    from openpyxl.worksheet.filters import AutoFilter
    for ws in wb.worksheets:
        ws._pivots=[]; ws.auto_filter=AutoFilter()
    names=wb.sheetnames
    P,C,I=f'{zen(prevmon)}月在庫',f'{zen(mon)}月仕入',f'{zen(mon)}月在庫'
    # 前月在庫データ（前月ファイルの当月在庫シート）
    prev=[(r[0],r[1],r[2]) for r in prev_inv_sheet.iter_rows(min_row=3,values_only=True) if r[0]]
    # シート名を当月に（テンプレの並び：前月在庫, 当月仕入, 当月在庫, 検証 ...）
    s_prev,s_cur,s_inv=wb.worksheets[0],wb.worksheets[1],wb.worksheets[2]
    s_inv.title='__tmp'; s_prev.title=P; s_cur.title=C; s_inv.title=I
    # 生データ差し替え
    for s in ['売上','仕入','kintone']: copy_raw(raw_wb[s], wb[s])
    for s in ['売上','仕入']:
        ws=wb[s]; last=ws.max_row
        for r in range(2,last+1):
            if ws.cell(r,1).value: ws.cell(r,25,f'=W{r}-V{r}' if s=='売上' else f'=V{r}-W{r}')
    kc=wb['kintoneコスト']; clear(kc)
    for row in raw_wb['kintoneコスト'].iter_rows(min_col=4,max_col=21):
        for c in row:
            if c.value is not None: kc.cell(c.row,c.column,c.value)
    # 計算（エンジン）
    _,sg=load_rows(raw_wb['売上']); _,pg=load_rows(raw_wb['仕入']); _,kt=load_rows(raw_wb['kintone'])
    res=compute(prev,sg,pg,kt,base_date)
    star,first=kintone_cost(raw_wb['kintoneコスト']); det=sales_detail(res,star,first,prev); mi=matome_inputs(det)
    # 売上PIVOT（値）
    sp=wb['売上PIVOT']
    # テンプレの売上明細の下にある集計ブロック（明細の最終行の後）を動的に探す
    _r=101
    while sp.cell(_r,1).value: _r+=1
    while _r<_r+20 and not any(sp.cell(_r,c).value is not None for c in range(1,11)) and _r<sp.max_row: _r+=1
    T0=_r; _e=_r
    while any(sp.cell(_e,c).value is not None for c in range(1,11)): _e+=1
    tail=[[c.value for c in row] for row in sp.iter_rows(min_row=T0,max_row=_e-1,max_col=10)]
    clear(sp)
    accts=['売上高','売上高（土地）']
    sp['A3'],sp['B3']='合計 / 差引額','列ラベル'
    for j,h in enumerate(['行ラベル']+accts+['総計'],1): sp.cell(4,j,h)
    memos=sorted((m for m in res['sales'] if m!='(空白)'),key=lambda m:(0 if str(m).startswith('案件') else 1,str(m))); r=5
    for m in memos:
        a=[res['sales_by_acct'][m].get(x,0) for x in accts]
        sp.cell(r,1,m); 
        for j,v in enumerate(a,2):
            if v: sp.cell(r,j,v)
        sp.cell(r,4,sum(a)); r+=1
    pv_last=r-1; tot=[sum(res['sales_by_acct'][m].get(x,0) for m in memos) for x in accts]
    sp.cell(r,1,'総計'); sp.cell(r,2,tot[0]); sp.cell(r,3,tot[1]); sp.cell(r,4,sum(tot))
    # 配置（先方ファイルと同じ）：ピボットの総計行の8行下に集計ブロック、その5行下に売上明細の見出し
    FR=max(95,r+8); H=FR+5
    n=len(det); d1,d2=H+1,H+n
    X=f'X{H}:X{d2}'
    sp[f'A{FR}'],sp[f'B{FR}'],sp[f'C{FR}'],sp[f'D{FR}']='freee',tot[0],tot[1],f'=SUM(B{FR}:C{FR})'
    sp[f'D{FR+1}']=f'=COUNTA(A5:A{pv_last})'; sp[f'E{FR+1}']='三為'; sp[f'F{FR+1}']=f'=COUNTIF(X{d1}:X{d2},0)'
    sp[f'G{FR+1}']=f'=SUMIFS(F{H}:F{d2},{X},0)/SUMIFS(C{H}:C{d2},{X},0)'; sp[f'H{FR+1}']=f'=SUMIFS(C{H}:C{d2},{X},0)'
    sp[f'D{FR+2}']=f'=COUNTA(A{d1}:A{d2})'; sp[f'E{FR+2}']='通常'; sp[f'F{FR+2}']=f'=D{FR+1}-F{FR+1}'
    sp[f'G{FR+2}']=f'=(SUM(F{H}:F{d2})-SUMIFS(F{H}:F{d2},{X},0))/(SUM(C{H}:C{d2})-SUMIFS(C{H}:C{d2},{X},0))'
    sp[f'H{FR+2}']=f'=(SUM(C{H}:C{d2})-SUMIFS(C{H}:C{d2},{X},0))'; sp[f'H{FR+3}']=f'=SUM(H{FR+1}:H{FR+2})=D{FR}'
    for col in 'CDEF': sp[f'{col}{FR+4}']=f'=SUM({col}{d1}:{col}{d2})'
    sp[f'G{FR+4}']=f'=F{FR+4}/C{FR+4}'
    hdr=['メモタグ','案件名','当月売上','前月在庫','当月仕入','A：freee粗利','B：kinton★コスト（管理費除く）','C：支払い漏れ/\n前期以前登記費用　などをfreee月次推移から抽出手入力\n火災保険','D：A-B-C\n\n最終粗利（freee+キントーンその他費用）','D/売上\n粗利率','E：キントーン粗利','キントーン粗利率','D-E\n㋐粗利差異','粗利率\n差異','①仮受消費税（カスタムレポートでBSメモタグ集計）','②仮払消費税（カスタムレポートでBSメモタグ集計）','①-②\n㋑消費税差額','㋐-㋑\n粗利差異（消費税調整後）','粗利率\n差異（消費税調整後）','リフォーム金額（万円）','コメント','仕入決済','売却決済','仕入から売却の月数']
    for j,h in enumerate(hdr,1): sp.cell(H,j,h)
    style_pivot(sp,pv_last,r,FR,H,d1,d2)
    kc_last=1+len(memos)
    for i,d in enumerate(det):
        r=d1+i; m=d['memo']
        sp[f'A{r}']=m
        sp[f'B{r}']=f'=IFERROR(IF(VLOOKUP(A{r},kintone!C:W,5,0)="","",VLOOKUP(A{r},kintone!C:W,5,0)),"")'
        sp[f'C{r}']=d['sales']
        sp[f'D{r}']=f"=IFERROR(VLOOKUP(A{r},'{P}'!$A$3:$C${len(prev)+2},3,0),0)"
        sp[f'E{r}']=f'=IFERROR(VLOOKUP(A{r},仕入PIVOT!$A$5:$K${5+len(res["purch"])},11,0),0)'
        sp[f'F{r}']=f'=C{r}-D{r}-E{r}'
        sp[f'G{r}']=f'=IFERROR(VLOOKUP(A{r},kintoneコスト!$A$2:$C${kc_last},3,0),0)'
        sp[f'I{r}']=f'=F{r}-G{r}-H{r}'; sp[f'J{r}']=f'=I{r}/C{r}*100'
        sp[f'K{r}']=f'=VLOOKUP(A{r},kintoneコスト!E:L,7,0)'; sp[f'L{r}']=f'=VLOOKUP(A{r},kintoneコスト!E:L,8,0)'
        sp[f'M{r}']=f'=I{r}-K{r}'; sp[f'N{r}']=f'=J{r}-L{r}'
        sp[f'Q{r}']=f'=O{r}-P{r}'; sp[f'R{r}']=f'=Q{r}+M{r}'; sp[f'S{r}']=f'=((I{r}+Q{r})-K{r})/C{r}*100'
        sp[f'V{r}']=f'=IFERROR(IF(VLOOKUP(A{r},kintone!C:W,14,0)="","",VLOOKUP(A{r},kintone!C:W,14,0)),"")'
        sp[f'W{r}']=f'=IFERROR(IF(VLOOKUP(A{r},kintone!C:W,19,0)="","",VLOOKUP(A{r},kintone!C:W,19,0)),"")'
        sp[f'X{r}']=f'=W{r}-V{r}'
        for col in 'CDEFGHIKMOPQR': sp[f'{col}{r}'].number_format='#,##0'
        for col in 'JLNSX': sp[f'{col}{r}'].number_format='General'
        for col in 'VW': sp[f'{col}{r}'].number_format='yyyy/mm/dd'
    for i,row in enumerate(tail):
        for j,v in enumerate(row,1):
            if v is not None:
                if isinstance(v,str) and v.startswith('='):
                    import re; v=re.sub(r'(?<=[A-Z])(\d+)(?!\d)',lambda mo:str(int(mo.group(1))-T0+d2+3) if T0<=int(mo.group(1))<T0+len(tail) else mo.group(1),v)
                sp.cell(d2+3+i,j,v)
    # kintoneコスト A〜C
    for i,m in enumerate(memos):
        r=2+i; kc[f'A{r}']=m; kc[f'B{r}']=f'=SUMIFS(S:S,E:E,A{r})-C{r}'; kc[f'C{r}']=f'=SUMIFS(S:S,E:E,A{r},N:N,"★")'
    kc['A1'],kc['B1'],kc['C1']='行ラベル','★以外コスト','★コスト'
    # 仕入PIVOT（値）
    pp=wb['仕入PIVOT']; clear(pp)
    pacc=['仕入高','仕入高（土地）','仕入高（報酬）','仕入高（報酬）土地','仕入高（租税公課）','仕入高（買取再販仲手）','仕入高（買取再販仲手）土地','仕入諸費用（土地建物）','仕入諸費用（土地）']
    pp['A3'],pp['B3']='合計 / 差引','列ラベル'
    for j,h in enumerate(['行ラベル']+pacc+['総計'],1): pp.cell(4,j,h)
    r=5; cols=[0]*len(pacc)
    for m,v in res['purch'].items():
        pp.cell(r,1,m)
        for j,a in enumerate(pacc):
            x=res['purch_by_acct'][m].get(a)
            if x is not None: pp.cell(r,2+j,x); cols[j]+=x
        pp.cell(r,11,v); r+=1
    pp.cell(r,1,'総計')
    for j,x in enumerate(cols): pp.cell(r,2+j,x)
    pp.cell(r,11,sum(cols)); pp_tot=r
    # 前月在庫シート
    sh=wb[P]; clear(sh,2)
    sh['D1']='棚卸用ソート条件：売上、仕入ともに0のセルを選択し前月在庫を棚卸へ'
    for j,h in enumerate(['メモタグ','案件名','前月在庫','当月売上','当月仕入','仕入決済日','売上決済日','買取再販以外','kintone進捗'],1): sh.cell(2,j,h)
    for i,(m,nm,a) in enumerate(prev):
        r=3+i; sh[f'A{r}']=m
        sh[f'B{r}']=f'=IFERROR(IF(VLOOKUP(A{r},kintone!C:W,5,0)="","",VLOOKUP(A{r},kintone!C:W,5,0)),"")'
        sh[f'C{r}']=a
        sh[f'D{r}']=f'=IFERROR(VLOOKUP(A{r},売上PIVOT!$A$5:$D${pv_last},4,0),0)'
        sh[f'E{r}']=f'=IFERROR(VLOOKUP(A{r},仕入PIVOT!$A$5:$K${pp_tot-1},11,0),0)'
        sh[f'F{r}']=f'=IFERROR(IF(VLOOKUP(A{r},kintone!C:W,14,0)="","",VLOOKUP(A{r},kintone!C:W,14,0)),"")'
        sh[f'G{r}']=f'=IFERROR(IF(VLOOKUP(A{r},kintone!C:W,19,0)="","",VLOOKUP(A{r},kintone!C:W,19,0)),"")'
        sh[f'J{r}']=f"=VLOOKUP(A{r},'{I}'!A:B,2,)='{P}'!B{r}"
    sh[f'C{3+len(prev)}']=f'=SUM(C3:C{2+len(prev)})'
    # 当月仕入シート
    sh=wb[C]; clear(sh)
    nc=len(res['cur_tbl']); last=3+nc
    sh['D1']=f'=SUM(D4:D{last})=仕入PIVOT!K{pp_tot}'; sh['E1']='台帳手入力'
    sh['F1']='※売上発生分の決済日が入力されているか確認し、売上0円の案件もチェックし間違えて決済日を入力していないか確認する※'
    sh['D2']='当月仕入0円は行削除済み'; sh['G2']=base_date; sh['H2']='棚卸用ソート条件：0以外を選択'
    for j,h in enumerate(['メモタグ','案件名','前月在庫','当月仕入高','仕入決済日','当月売上　','売上決済日','当月在庫','kintone進捗'],1): sh.cell(3,j,h)
    for i,t in enumerate(res['cur_tbl']):
        r=4+i; sh[f'A{r}']=t['memo']
        sh[f'B{r}']=f'=IFERROR(IF(VLOOKUP(A{r},kintone!C:W,5,0)="","",VLOOKUP(A{r},kintone!C:W,5,0)),"")'
        sh[f'C{r}']=f"=IFERROR(VLOOKUP(A{r},'{P}'!$A$3:$C${len(prev)+2},3,0),0)"
        sh[f'D{r}']=t['purch']
        sh[f'E{r}']=f'=IFERROR(IF(VLOOKUP(A{r},kintone!C:W,14,0)="","",VLOOKUP(A{r},kintone!C:W,14,0)),"")'
        sh[f'F{r}']=f'=IFERROR(VLOOKUP(A{r},売上PIVOT!$A$5:$D${pv_last},4,0),0)'
        sh[f'G{r}']=f'=IFERROR(IF(VLOOKUP(A{r},kintone!C:W,19,0)="","",VLOOKUP(A{r},kintone!C:W,19,0)),"")'
        sh[f'H{r}']=f'=IF(OR(G{r}="",G{r}>$G$2),C{r}+D{r},0)'
        sh[f'I{r}']=f'=IFERROR(IF(VLOOKUP(A{r},kintone!A:W,6,0)="","",VLOOKUP(A{r},kintone!A:W,6,0)),"")'
        sh[f'J{r}']=f"=VLOOKUP(A{r},'{P}'!A:A,1,0)"
    sh[f'C{last+1}']=f'=SUM(C4:C{last})'; sh[f'D{last+1}']=f'=SUM(D4:D{last})'
    # 当月在庫シート
    sh=wb[I]; clear(sh); inv=res['inv']; li=2+len(inv)
    sh['A1']=f'=COUNTA(A3:A{li})'; sh['C1']=f'=SUM(C3:C{li})'; sh['D1']='基準日'; sh['E1']=f"='{C}'!G2"
    for j,h in enumerate(['メモタグ','摘要','金額','仕入決済','滞留日数','ダブリチェック'],1): sh.cell(2,j,h)
    for j,h in zip('HIJK',['作成者','作成日','承認者','承認日']): sh[f'{j}2']=h
    for i,x in enumerate(inv):
        r=3+i; sh[f'A{r}']=x['memo']; sh[f'B{r}']=x['name']; sh[f'C{r}']=x['amount']
        sh[f'D{r}']=f'=IFERROR(IF(VLOOKUP(A{r},kintone!C:W,14,0)="","",VLOOKUP(A{r},kintone!C:W,14,0)),"")'
        sh[f'E{r}']=f'=IF(D{r}="","未決済",DATEDIF(D{r},$E$1,"D"))'; sh[f'F{r}']=f'=COUNTIF(A:A,A{r})'
    # 検証
    sh=wb['検証']; clear(sh)
    for j,h in enumerate(['メモタグ','案件名','棚卸金額','前月在庫','仕入','合計','差異','重複チェック','売上チェック'],1): sh.cell(1,j,h)
    for i,x in enumerate(inv):
        r=2+i; sh[f'A{r}']=x['memo']; sh[f'B{r}']=x['name']; sh[f'C{r}']=x['amount']
        sh[f'D{r}']=f"=IFERROR(VLOOKUP(A{r},'{P}'!$A$3:$C${len(prev)+2},3,0),0)"
        sh[f'E{r}']=f'=IFERROR(VLOOKUP(A{r},仕入PIVOT!$A$5:$K${pp_tot-1},11,0),0)'
        sh[f'F{r}']=f'=+D{r}+E{r}'; sh[f'G{r}']=f'=C{r}-F{r}'
        sh[f'H{r}']=f'=COUNTIF($A$2:$A${1+len(inv)},A{r})'
        sh[f'I{r}']=f'=IFERROR(VLOOKUP(A{r},売上PIVOT!$A$2:$D${pv_last},4,0),0)'
    # 要確認
    if 'AI要確認' in wb.sheetnames: del wb['AI要確認']
    ws=wb.create_sheet('AI要確認',0)
    ws['A1']=f'{mon}月棚卸：AIが自動作成。人の判断が必要な項目'; ws['A1'].font=Font(bold=True)
    for j,h in enumerate(['内容','メモタグ','金額'],1): ws.cell(3,j,h).font=Font(bold=True)
    rows=res['alerts']+res['excluded']
    for i,a in enumerate(rows):
        for j,v in enumerate(a,1): ws.cell(4+i,j,v)
    ws.column_dimensions['A'].width=70; ws.column_dimensions['B'].width=18; ws.column_dimensions['C'].width=16
    wb.save(out)
    return res,det,mi
if __name__=='__main__':
    raw=openpyxl.load_workbook('kai09.xlsx',data_only=True)
    k8=openpyxl.load_workbook('kai08.xlsx',data_only=True,read_only=True)
    res,det,mi=build('kai08.xlsx','kai09_ai_raw.xlsx',raw,k8['８月在庫'],9,8,datetime.datetime(2026,9,30))
    print(len(res['inv']),sum(x['amount'] for x in res['inv']),mi)
