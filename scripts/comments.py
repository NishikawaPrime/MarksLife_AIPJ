"""PL増減分析の「数値から機械的に書けるコメント」を生成（BC=単月、BH=累計）"""
def K(x): return f'{x:,.0f}K'
def pct(x): return f'{x*100:.1f}%'
def gen(mt, pl, mon, fudosan=True):
    """mt(col,row)->value from 買取再販まとめ; pl(label,col)->value from PL増減分析"""
    out={}
    for scope,(pc,cc,bx) in {'単月':(mt.cur_prev,mt.cur,'BX'),'累計':(mt.cum_prev,mt.cum,'BZ')}.items():
        g=lambda col,row: mt.v(col,row) or 0
        n0,n1=g(pc,12),g(cc,12); s0,s1,u0,u1=g(pc,10),g(cc,10),g(pc,11),g(cc,11)
        p0=int(s0/n0*100); p1=int(s1/n1*100)   # 同日は切捨て、通常は100-同日（既存資料の書き方に合わせる）
        L=['件数',
           f'  前期：{n0:.0f}件（同日 {s0:.0f}件〈{p0}％〉、通常 {u0:.0f}件〈{100-p0}％〉）',
           f'  当期：{n1:.0f}件（同日 {s1:.0f}件〈{p1}％〉、通常 {u1:.0f}件〈{100-p1}％〉）',
           '1件当たり売上高',
           f'  前期：{K(g(pc,42))}（同日 {K(g(pc,40))}、通常 {K(g(pc,41))}）',
           f'  当期：{K(g(cc,42))}（同日 {K(g(cc,40))}、通常 {K(g(cc,41))}）',
           '1人当たり売上高',
           f'  前期：{K(g(pc,32))}（同日 {K(g(pc,30))}、通常 {K(g(pc,31))}）',
           f'  当期：{K(g(cc,32))}（同日 {K(g(cc,30))}、通常 {K(g(cc,31))}）']
        out[('買取再販合計',scope)]='\n'.join(L)
        L=['1件当たり粗利',
           f'  前期：{K(g(pc,47))}（同日 {K(g(pc,45))}、通常 {K(g(pc,46))}）',
           f'  当期：{K(g(cc,47))}（同日 {K(g(cc,45))}、通常 {K(g(cc,46))}）']
        if fudosan: L.append(f'            {K(g(bx,47))}（同日 {K(g(bx,45))}、通常 {K(g(bx,46))}）・・・ 不動産取得税振替前')
        L+=['1人当たり粗利',
           f'  前期：{K(g(pc,37))}（同日 {K(g(pc,35))}、通常 {K(g(pc,36))}）',
           f'  当期：{K(g(cc,37))}（同日 {K(g(cc,35))}、通常 {K(g(cc,36))}）']
        if fudosan: L.append(f'            {K(g(bx,37))}（同日 {K(g(bx,35))}、通常 {K(g(bx,36))}）・・・ 不動産取得税振替前')
        out[('買取再販粗利',scope)]='\n'.join(L)
        L=[f'  前期：同日 {pct(g(pc,25))}、通常 {pct(g(pc,26))}',f'  当期：同日 {pct(g(cc,25))}、通常 {pct(g(cc,26))}']
        if fudosan: L.append(f'            同日 {pct(g(bx,25))}、通常 {pct(g(bx,26))} ・・・ 不動産取得税振替前')
        out[('粗利率（買取再販）',scope)]='\n'.join(L)
        i=0 if scope=='単月' else 1
        r=lambda lab,prev: pl(lab,scope,prev)
        out[('仲介手数料',scope)]=f"比率16期：{r('仲介手数料',1)/r('売上高 計',1)*100:.1f}％　17期：{r('仲介手数料',0)/r('売上高 計',0)*100:.1f}％"
        base=lambda p: sum(r(x,p) for x in ('役員報酬','給料手当','賞与','雑給','賞与引当金繰入'))
        out[('法定福利費',scope)]=f"比率（役員報酬、給料、賞与、雑給、賞与引当金繰入に対する比率）\n16期：{r('法定福利費',1)/base(1)*100:.1f}％　17期：{r('法定福利費',0)/base(0)*100:.1f}％"
    return out
