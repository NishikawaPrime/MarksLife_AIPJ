import csv,glob,json,sys,subprocess,unicodedata,re
sys.path.insert(0,'.')
from build_month import build
import manual
D='/home/claude/oct/in/'
def rcsv(p):
    fs=[f for f in glob.glob(D+p) if 'Ｍ＆Ａ' not in f]; assert len(fs)==1,fs
    return list(csv.reader(open(fs[0],encoding='cp932').read().splitlines()))
def conv(v):
    if v in ('',None): return None
    try: return int(v)
    except:
        try: return float(v)
        except: return v
BS=[[conv(x) for x in r]+[None]*(19-len(r)) for r in rcsv('f3.csv')]
PL=[[conv(x) for x in r[:18]] for r in rcsv('f4.csv')]
BS=[r[:19] for r in BS]
def cut(block,first_month_col,n):
    """n月目より後の列を空にする"""
    out=[]
    for i,r in enumerate(block):
        r=list(r)
        if i>=2:
            for j in range(first_month_col+n,len(r)): r[j]=None
        out.append(r)
    return out
def mats(n):
    base=json.load(open('materials_10.json'))
    m=dict(month_index=n,month_label=manual.MONTHS[n-1],prev_month_label=manual.MONTHS[n-2],
           bs_title=f'2026年-{manual.MONTHS[n-1][:-1]}月期末',
           freee_bs=cut(BS,7,n), freee_pl=cut(PL,6,n), manual=manual.values(n),
           fudosan_k=base['fudosan_k'], matome=base['matome'],
           sales_pivot=base['sales_pivot'], sales_detail=base['sales_detail'])
    return m
HW=re.compile(r'[｡-ﾟ]+')
def run(cmd): 
    r=subprocess.run(cmd,capture_output=True,text=True); print(r.stdout.strip()[-400:],r.stderr.strip()[-400:]); assert r.returncode==0
def month(src,tag,n,mode,check,manual_notes):
    m=mats(n); mon=m['month_label']
    log=build(src,f'{tag}_1.xlsx',m,mode)
    cur=f'{tag}_1.xlsx'
    if n==12:
        import os; os.environ['PL_ADDS']='受贈益:固定資産売却益'
        run(['python3','patch_pl.py',cur,f'{tag}_2.xlsx']); cur=f'{tag}_2.xlsx'
        log.append('freeeに新しく出てきたPL科目「受贈益」を特別利益の区分に追加（マスタにも追加）')
    run(['python3','apply_comments.py',cur,str(n),mon,f'{tag}_gen.json'])
    g=json.load(open(f'{tag}_gen.json'))
    s,c=manual.ma_comment(n); g['MA会社費用|単月']=s; g['MA会社費用|累計']=c
    json.dump(g,open(f'{tag}_gen.json','w'),ensure_ascii=False)
    log+=['PL増減分析：買取再販合計・買取再販粗利・粗利率・仲介手数料・法定福利費・MA会社費用の差異コメント（単月／累計）を数値から作成']
    mv=m['manual']; i=n-1
    man={'外注費(派遣費用)':f"{mv['haken'][i]:,}円（本体の仕訳帳、品目「【外注費】人材派遣」）",
         '不動産取得税':f"{mv['fudosan'][i]:,}円（本体の仕訳帳、仕入高（租税公課）の品目「不動産取得税（確定）」「（概算）」）",
         'MA会社費用':f"{mv['ma'][i]:,}円（子会社マークスライフＭ＆Ａ株式会社の月次推移の販売管理費 計）",
         '営業人員':f"{mv['jinin'][i]}人（17期_人員 まとめの営業・社員・当年）"}
    man.update(manual_notes)
    json.dump(dict(done=log,manual=man,check=check),open(f'{tag}_log.json','w'),ensure_ascii=False)
    run(['python3','finalize.py',cur,f'{tag}_3.xlsx',str(n),mon,f'{tag}_gen.json',f'{tag}_log.json'])
    return f'{tag}_3.xlsx'
if __name__=='__main__':
    print(month('n09_4.xlsx','n10',12,'roll',
      ['10月は期末月。当月の法人税等は6,667円のみで、期末の法人税等の計上は未了の可能性が高い（freeeの月次推移どおりに作成。組替(BS)の未払法人税等は35%概算の累計式のまま）',
       '10月の当月純損益は大きくプラス（税引前591百万円）。特別利益「受贈益」53,854,685円、寄付金5,000,000円、賞与引当金繰入113,725,006円などの決算整理仕訳が10月31日付で304件入っている',
       'MA会社費用の10月分は子会社の月次推移PL（販売管理費 計2,640,760円）から。内訳は子会社の仕訳帳で確認（外注費は神崎信也氏994,500円。スマート・ストラテジー400,000円は本体（マークスライフ株式会社）への付替で相殺）。仕訳帳の旅費交通費は月次推移より818円少ない',
       '子会社の10月の営業外：雑損失474,157円は子会社の設立費用（本体へ支払）、支払利息139,315円は本体からの借入金3,000万円の利息（7/10〜10/31）。本体の10月BSに長期貸付金30,000,000円が新しく計上されている。本体側の受取利息・設立費用の受取りの処理は要確認。MA会社費用には含めていない',
       '棚卸：情報2025063776（名古屋市西区江向町 土地）は10月に2分割され（②を売却）、先方の10月棚卸の「9月在庫」に分割後の値が入っているため、それを前月在庫として使用',
       '棚卸：案件0004742（四條畷市田原台）の売上明細で、先方は前月在庫に+12,408,839円を手入力で加算している（根拠の資料なし）。AIは加算していないため、10月のfreee粗利合計は先方より12,408,839円大きい',
       '棚卸：案件0005346（高崎市飯塚町 土地）はfreeeで10月に売上計上されているが、kintoneに売却決済日がない（売却契約済のまま）。在庫からは除外し、日数は空欄'],{}))
