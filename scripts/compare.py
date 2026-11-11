import openpyxl,sys,collections
A=openpyxl.load_workbook(sys.argv[1],data_only=True); B=openpyxl.load_workbook(sys.argv[2],data_only=True)
specs={'月次推移':(6,list(range(7,23))),'BS増減分析':(1,[2,3,4]),'組替(BS)':(1,list(range(3,16))),
       'PL増減分析':(6,list(range(7,60))),'組替(PL)当年':(6,list(range(7,20))),'買取再販まとめ':(3,list(range(4,80)))}
def keyed(ws,lc):
    d={}; cnt=collections.Counter()
    for r in range(1,ws.max_row+1):
        lab=None
        for c in range(1,lc+1):
            v=ws.cell(r,c).value
            if v not in (None,''): lab=str(v).strip(); break
        if lab is None: lab=f'#row{r}'
        cnt[lab]+=1; d[(lab,cnt[lab])]=r
    return d
tot_ok=tot_ng=0
for s,(lc,cols) in specs.items():
    wa,wb=A[s],B[s]; ka,kb=keyed(wa,lc),keyed(wb,lc)
    ok=ng=0; ex=[]; miss=[k for k in kb if k not in ka and not k[0].startswith('#')]
    for k,rb in kb.items():
        if k not in ka: continue
        ra=ka[k]
        for c in cols:
            va,vb=wa.cell(ra,c).value,wb.cell(rb,c).value
            if isinstance(vb,(int,float)) and not isinstance(vb,bool):
                va2=va if isinstance(va,(int,float)) else 0
                if abs(va2-vb)<=max(1,abs(vb)*1e-9): ok+=1
                else:
                    ng+=1
                    if len(ex)<int(sys.argv[3]) : ex.append((k[0],wb.cell(rb,c).coordinate,va,vb))
    tot_ok+=ok; tot_ng+=ng
    print(f'== {s}: 一致 {ok} / 不一致 {ng}   実物にあって再現にない行: {miss[:12]}')
    for e in ex: print('    ',e)
print('TOTAL',tot_ok,tot_ng)
