import pandas as pd, re, openpyxl, json
L=pd.read_pickle('/home/claude/oct/c/L.pkl')
LP=pd.read_pickle('/home/claude/oct/c/LP.pkl')  # 前期（16期）仕訳
wv=openpyxl.load_workbook('/home/claude/oct/c/out.xlsx',data_only=True)
pl=wv['PL増減分析']; mt=wv['買取再販まとめ']
def k(x): return ('△' if x<0 else '')+f"{abs(x)/1000:,.0f}K"
def sk(x): return ('+' if x>=0 else '△')+f"{abs(x)/1000:,.0f}K"
def key(r):
    p=str(r['partner'] or '').strip(); it=re.sub(r'【[^】]*】','',str(r['item'] or '')).strip(); n=str(r['note'] or '').strip()
    n=re.sub(r'https?://\S+','',n).strip()
    if it=='その他': it=''
    import unicodedata
    return unicodedata.normalize('NFKC',p or it or n[:18] or '（摘要なし）')
REV={'業託・紹介料収入','家賃収入','雑収入','受贈益','仲介手数料','その他収入'}
def breakdown(accts,yms,top=4):
    d=L[L.acct.isin(accts)&L.ym.isin(yms)].copy()
    if d.empty: return None,0
    sg=-1 if accts[0] in REV else 1
    d['k']=d.apply(key,axis=1); g=(d.groupby('k').amt.sum()*sg).sort_values(key=abs,ascending=False)
    g=g[g.abs()>=1000]; tot=g.sum(); parts=[f"{i} {k(v)}" for i,v in g.head(top).items()]
    rest=tot-g.head(top).sum()
    if abs(rest)>=1000: parts.append(f"その他 {k(rest)}")
    return '、'.join(parts),g
def mom(accts):
    sg=-1 if accts[0] in REV else 1
    d=L[L.acct.isin(accts)&L.ym.isin(['2026-09','2026-10'])].copy(); d['k']=d.apply(key,axis=1)
    p=d.pivot_table(index='k',columns='ym',values='amt',aggfunc='sum',fill_value=0)*sg
    for c in ['2026-09','2026-10']:
        if c not in p: p[c]=0
    p['d']=p['2026-10']-p['2026-09']; t=p.d.sum(); p=p.reindex(p.d.abs().sort_values(ascending=False).index)
    big=[f"{i} {sk(v)}" for i,v in p.d.head(2).items() if abs(v)>=100000]
    return f"前月（9月）比 {sk(t)}"+(f"（主に{'、'.join(big)}）" if big else '')
def monthly_avg(accts):
    sg=-1 if accts[0] in REV else 1
    m=(L[L.acct.isin(accts)].groupby('ym').amt.sum()*sg).reindex(sorted(L.ym.unique()),fill_value=0)
    return f"月平均 {k(m.mean())}（8〜10月平均 {k(m.loc[['2026-08','2026-09','2026-10']].mean())}）"
CNT=lambda col: {'同日':mt[f'{col}10'].value,'通常':mt[f'{col}11'].value,'計':mt[f'{col}12'].value,'売上':mt[f'{col}17'].value}
def sales(single):
    a,b=('AK','AL') if single else ('BS','BT')
    p,c=CNT(a),CNT(b)
    s=f"件数 前期{p['計']}件→当期{c['計']}件（同日{p['同日']}→{c['同日']}件、通常{p['通常']}→{c['通常']}件）\n1件当たり売上 前期{p['売上']/p['計']:,.0f}K→当期{c['売上']/c['計']:,.0f}K"
    if single:
        sm=wv['売上明細(10月)']; rows=[(r[0],r[1],r[2]) for r in sm.iter_rows(min_row=101,values_only=True) if isinstance(r[2],(int,float)) and r[0] and re.match(r'(案件|情報|買取再販)',str(r[0]))]
        rows.sort(key=lambda x:-x[2]); s+="\n大型案件："+'、'.join(f"{(n or t)} {k(v)}" for t,n,v in rows[:3])
    return s
ACC={'業託・紹介料収入':['業託・紹介料収入'],'家賃収入':['家賃収入'],'役員報酬':['役員報酬'],'給料手当':['給料手当'],'雑給':['雑給'],'福利厚生費':['福利厚生費'],
 '採用費(前年:採用教育費)':['採用費'],'外注費':['外注費'],'広告宣伝費':['広告宣伝費'],'支払手数料':['支払手数料'],'支払報酬料':['支払報酬料'],'消耗品費':['消耗品費'],
 '保険料':['保険料'],'租税公課':['租税公課'],'減価償却費':['減価償却費'],'旅費交通費':['旅費交通費'],'通信費':['通信費'],'交際費':['交際費'],'会議費':['会議費'],
 '地代家賃':['地代家賃'],'リース料':['リース料'],'寄付金':['寄付金'],'履行差額（販管費）':['履行差額（販管費）'],'雑収入':['雑収入'],'支払利息':['支払利息'],
 '融資手数料':['融資手数料'],'受贈益':['受贈益']}
def lab(r): return next((pl.cell(r,c).value for c in (1,2,3,4) if pl.cell(r,c).value),'')
TAIL='増減の理由：要記入'
def make(r,single):
    l=lab(r); yms=['2026-10'] if single else sorted(L.ym.unique())
    p,c,dd=(pl.cell(r,51).value,pl.cell(r,52).value,pl.cell(r,53).value) if single else (pl.cell(r,56).value,pl.cell(r,57).value,pl.cell(r,58).value)
    head=f"前期{k(p)}→当期{k(c)}（{sk(dd)}）"
    if l in ('売上高','売上高（土地）'): return head+'\n'+sales(single)+'\n'+('売上高＝建物部分、売上高（土地）＝土地部分（件数は合計）')
    if l=='不動産取得税':
        d=L[(L.acct=='租税公課')&L.ym.isin(yms)&L.item.fillna('').str.contains('不動産取得税')]
        pv=LP[(LP.acct=='租税公課')&(LP.ym.isin(['2025-10']) if single else True)&LP.item.str.contains('不動産取得税')].amt.sum()
        return head+f"\n当期から不動産取得税を租税公課（販管費）から売上原価へ振替（{'当月' if single else '累計'} {k(c)}）。前期は租税公課に計上（{'前期10月' if single else '前期累計'} {k(pv)}）"
    if l=='商品売上原価':
        sp,sc=(pl.cell(8,51).value,pl.cell(8,52).value) if single else (pl.cell(8,56).value,pl.cell(8,57).value)
        return head+f"\n売上（買取再販）{sk((sc-sp))}に伴う増加。原価率 前期{p/sp:.1%}→当期{c/sc:.1%}"
    if l=='買取再販粗利（不動産取得税振替前）':
        sp,sc=(pl.cell(8,51).value,pl.cell(8,52).value) if single else (pl.cell(8,56).value,pl.cell(8,57).value)
        return head+f"\n粗利率 前期{p/sp:.1%}→当期{c/sc:.1%}（不動産取得税を原価に含めない場合）"
    if l=='賞与引当金繰入':
        x=f"\n当月：決算引当金平準化+117,063K（インセンティブ108,907K、BO賞与8,155K）、感謝手当戻入△3,337K" if single else "\n期首の引当金取崩し（11月△224,463K）後、毎月の平準化繰入で積み増し"
        return head+x+'\n'+yoy(['賞与引当金繰入'],single)+'\n'+TAIL
    if l=='労務費':
        kids=[(lab(i),(pl.cell(i,53) if single else pl.cell(i,58)).value) for i in range(53,62)]
        kids=sorted([x for x in kids if isinstance(x[1],(int,float))],key=lambda x:-abs(x[1]))[:3]
        return head+"\n主な増減："+'、'.join(f"{a} {sk(b)}" for a,b in kids)+(f"\n営業人員 前期{mt['AK5'].value}人→当期{mt['AL5'].value}人" if single else f"\n営業人員（平均）前期{mt['BS5'].value:.0f}人→当期{mt['BT5'].value:.0f}人")
    if l in ('経費','営業外費用','特別利益','特別損失'):
        rng={'経費':range(63,89),'営業外費用':range(100,107),'特別利益':range(109,114),'特別損失':range(115,120)}[l]
        kids=[(lab(i),(pl.cell(i,53) if single else pl.cell(i,58)).value) for i in rng]
        kids=sorted([x for x in kids if isinstance(x[1],(int,float)) and x[1]],key=lambda x:-abs(x[1]))[:4]
        return head+"\n主な増減："+'、'.join(f"{a} {sk(b)}" for a,b in kids)
    if l=='投資有価証券評価損': return head+"\n前期のみ計上（当期は計上なし）"
    if l=='法人税等' and c:
        f=L[(L.acct=='法人税等')&L.ym.isin(yms)].amt.sum()
        return head+f"\nfreeeの計上は{k(f)}（預金利子等の源泉税）。残りは月次報告で入れた税金の概算（期末の税金はfreee未計上）。確定額：要記入"
    if l=='法人税等調整額':
        return head+("\n当期10月は計上なし（前期10月は期末の税効果△72,167K）。期末の税効果の計上：要記入" if single else "\n当期は4月に△16,355K（中間決算時の税効果）。期末の税効果の計上：要記入")
    if l in ACC:
        bd,_=breakdown(ACC[l],yms)
        if bd is None: return head+"\n当期の仕訳なし"+'\n'+TAIL
        s=head+(f"\n当期の主な内訳：{bd}" if len(_)>1 else '')
        if l=='受贈益': return head+"\n10月：リブゼ横浜サザンティアラ201号室（土地19,466K・建物12,374K）の受贈 31,840K、ヤンマービル（本社）の内装・設備の譲渡 22,015K（贈与元・経緯：要記入）"
        s+='\n'+(mom(ACC[l]) if single else monthly_avg(ACC[l]))
        if l=='役員報酬': s+='\n'+('月額 前期5,700K→当期7,650K（改定・増員の内容：要記入）')
        if l=='給料手当': s+='\n'+(f"営業人員 前期{mt['AK5'].value}人→当期{mt['AL5'].value}人" if single else f"営業人員（平均）前期{mt['BS5'].value:.0f}人→当期{mt['BT5'].value:.0f}人")
        if l=='租税公課': s+='\n'+'※当期は不動産取得税を売上原価へ振替（前期は租税公課に含む）'
        return s+'\n'+yoy(ACC[l],single)+'\n'+TAIL
    return None

def yoy(accts,single):
    sg=-1 if accts[0] in REV else 1
    a=L[L.acct.isin(accts)&(L.ym.isin(['2026-10']) if single else True)].copy()
    b=LP[LP.acct.isin(accts)&(LP.ym.isin(['2025-10']) if single else True)].copy()
    for x in (a,b): x['k']=x.apply(key,axis=1) if len(x) else []
    ga=a.groupby('k').amt.sum()*sg if len(a) else pd.Series(dtype=float)
    gb=b.groupby('k').amt.sum()*sg if len(b) else pd.Series(dtype=float)
    dd=ga.sub(gb,fill_value=0); dd=dd[dd.abs()>=100000].reindex(dd.abs().sort_values(ascending=False).index).dropna()
    top=[f"{i} {sk(v)}" for i,v in dd.head(4).items()]
    gb=gb[gb.abs()>=1000].reindex(gb.abs().sort_values(ascending=False).index).dropna()
    pre='前期の主な内訳：'+'、'.join(f"{i} {k(v)}" for i,v in gb.head(3).items()) if len(gb) else '前期の計上なし'
    return pre+('\n前期比の主な増減：'+'、'.join(top) if top else '')
