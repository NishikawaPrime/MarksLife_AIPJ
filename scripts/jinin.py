import openpyxl
def eigyo_shain(path,n):
    ws=openpyxl.load_workbook(path,data_only=True)['まとめ']
    c=8+5*(n-1)+2
    tot=0
    for r in ws.iter_rows(min_row=1,values_only=True):
        if r[5]=='社員' and r[7]=='営業' and isinstance(r[c],(int,float)): tot+=r[c]
    return tot
if __name__=='__main__':
    for n in (9,10,11): print(n,eigyo_shain('/mnt/user-data/uploads/Downloads/17期_人員 1.xlsx',n))
