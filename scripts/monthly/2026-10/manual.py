"""手入力値（freee本体以外から反映する数値）の月別データ。index 1=11月 ... 12=10月"""
import json,sys
sys.path.insert(0,'.')
from jinin import eigyo_shain
MONTHS=['11月','12月','1月','2月','3月','4月','5月','6月','7月','8月','9月','10月']
YM=['2025-11','2025-12','2026-01','2026-02','2026-03','2026-04','2026-05','2026-06','2026-07','2026-08','2026-09','2026-10']
# 本体の仕訳帳（2025/05〜2026/09、再送分）から集計
HAKEN=[1307442,1596072,1103227,1664150,1535482,1557588,1454404,1523448,1550888,1406937,1449876,1525304]
FUDOSAN={6:53775000+28623500,7:5655800+7695400,8:9069300+5739600,9:4347900+6874000,10:-1524200+10032500,11:3251100+8695300,12:11853800}
# 子会社マークスライフＭ＆Ａ株式会社の月次推移（PL）の販売管理費 計
MA={9:2968665,10:5613294,11:3772297,12:2640760}
JININ_SRC='/home/claude/oct/in/f6.xlsx'
def values(n):
    """n月目までの値（それ以降は空）"""
    v={}
    v['haken']=[HAKEN[i] if i<n else None for i in range(12)]
    v['fudosan']=[(FUDOSAN.get(i+1,0) if i<n else None) for i in range(12)]
    v['ma']=[(MA.get(i+1,0) if i<n else None) for i in range(12)]
    v['ma_accr_month']=11   # MAの累計を未払費用に計上し始めた月（9月の取締役会資料から）
    v['jinin']=[eigyo_shain(JININ_SRC,i+1) if i<n else None for i in range(12)]
    return v
def ma_comment(n):
    d=json.load(open('ma_detail.json'))
    ym=YM[n-1]; yms=[y for y in YM[:n]]
    k=lambda x:f'{int(x/1000):,}K'
    order=['役員報酬・法定福利費','神崎信也 業務委託','スマート・ストラテジー（M&Aサービス開発）','マークスライフ株式会社への経営指導料','広告宣伝費','その他']
    def line(get):
        parts=[f'{c}+{k(get(c))}' for c in order if c in d and get(c)]
        return '、'.join(parts)
    s=line(lambda c:d[c].get(ym,0))
    c=line(lambda c:sum(d[c].get(y,0) for y in yms))
    head='子会社マークスライフＭ＆Ａ株式会社の販売管理費\n'
    return head+s, head+c
if __name__=='__main__':
    print(values(11)); print(ma_comment(10)); print(ma_comment(11))
