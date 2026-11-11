"""取締役会資料2枚目（BS）の数字を月次報告のBS増減分析から作る"""
import openpyxl,math,sys
sys.path.insert(0,'.')
from board_vals import vals
def tz(x): return int(x/1e6) if x>=0 else -int(-x/1e6)
def sd(v): return f'+{v:,}' if v>=0 else f'△{abs(v):,}'
def bs_params(report,n,mon,inv_count,kishu_count=203,kishu_avg='10,025K'):
    ws=openpyxl.load_workbook(report,data_only=True)['BS増減分析']
    B={};C={}; lab=[]
    for r in range(1,ws.max_row+1):
        l=ws.cell(r,1).value
        if l is None: continue
        l=str(l).strip(); b,c=ws.cell(r,2).value,ws.cell(r,3).value
        if l not in B: B[l]=b if isinstance(b,(int,float)) else 0; C[l]=c if isinstance(c,(int,float)) else 0; lab.append(l)
    d=lambda *ks: sum(C[k]-B[k] for k in ks)
    def side(X):
        total=tz(X['負債及び純資産 計']); cash=tz(X['現預金']); inv=tz(X['在庫']); fixed=tz(X['固定資産 計'])
        loan=tz(X['借入金']); eq=tz(X['純資産 計'])
        return dict(total=total,cash=cash,inv=inv,fixed=fixed,other=total-cash-inv-fixed,loan=loan,eq=eq,oliab=total-loan-eq)
    k,c=side(B),side(C)
    chart=dict(other=[k['other'],c['other'],k['eq'],c['eq']],fixed=[k['fixed'],c['fixed'],k['oliab'],c['oliab']],
               inv=[k['inv'],c['inv'],k['loan'],c['loan']],cash=[k['cash'],c['cash'],None,None],total=[k['total'],c['total'],k['total'],c['total']])
    V=vals(report,n); net=tz(V['当期純損益金額']['c1'])
    dc=(C['現預金']-B['現預金'])/1e8; dc=math.trunc(dc*10)/10
    ratio=lambda X: X['純資産 計']/X['負債及び純資産 計']*100
    tsu=sum(C[l] for l in lab if '（通知預金）' in l); tei=sum(C[l] for l in lab if '（定期積金）' in l or '（定期預金）' in l)
    avg=int(C['在庫']/inv_count/1000)
    p=dict(chart=chart,
      cash_lines=[f"・現預金{'△' if dc<0 else '＋'}{abs(dc):.1f}億",'　▸内訳（営業活動・新店投資等・在庫増加ーPJ借入・','　　税金賞与融資の返済・運転資金社債）は要確認'],
      inv_lines=['【期首】',f'    件数：{kishu_count}件',f'    在庫@：{kishu_avg}','',f'【{mon}末】',f'    件数：{inv_count}件',f'    在庫＠：{avg:,}K'],
      eq_lines=[f'◇当期純利益{sd(net) if net<0 else net}','◇自己資本比率',f'  期  首：{ratio(B):.1f}%',f'  {mon}末：{ratio(C):.1f}%'],
      liab_lines=[f"◇納税 {sd(tz(d('未払法人税等','未払事業税等','未払消費税等')))}",f"◇賞与引当金 {sd(tz(d('賞与引当金')))}",
                  f"◇未払給与 {sd(tz(d('未払給与')))}",f"◇未払費用 {sd(tz(d('未払費用')))}",f"◇手付金 {sd(tz(d('前受金')))}"],liab_red=(),
      oth_lines=[f"未収還付消費税{sd(tz(d('未収還付消費税等')))}、手付金{sd(tz(d('前渡金')))}",f"前払費用(支払利息・他){sd(tz(d('前払費用')))}"],
      loan_lines=[f"【借入 純増{sd(tz(d('借入金')))}】",'◇PJ資金・運転資金・社債の','  借入／返済の内訳は要確認'],
      kintone_lines=['（在庫粗利：要確認）','Kintone成約予定ベース'],
      dep_lines=[f'通知預金：{tz(tsu)}M',f'定期積金：{tz(tei)}M'],
      fixed_lines=[f"支店開設敷金{sd(tz(d('敷金')))}",f"附属設備{sd(tz(d('附属設備')))}",f"豊岡出石{sd(tz(d('土地','建設仮勘定')))}"])
    return p
