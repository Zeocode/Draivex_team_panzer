import xlrd
loc = "E:/Important/Project/Draivex/test data/dataset.xlsx"
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
i=0
j=0
row=1
while(row<sheet.nrows):
    if ((sheet.cell_value(row, 4)=="ME") or(sheet.cell_value(row, 4)=="CE")):
        if sheet.cell_value(row, 3)==("3rd" or "4th"):
            x=sheet.cell_value(row, 2)
            y=sheet.cell_value(row,1)
            fn=(y.split(" ")[0])
            ln=(y.split(" ")[1])
            if ((fn.lower() in x.lower()) or (ln.lower() in x.lower())):
                print("",end="")
            else:
                i+=1
        else:
            i+=1
    else:
        if sheet.cell_value(row, 3)==("3rd" or "4th"):
            x=sheet.cell_value(row, 2)
            y=sheet.cell_value(row,1)
            fn=(y.split(" ")[0])
            ln=(y.split(" ")[1])
            if (fn.lower() in x.lower()) or (ln.lower() in x.lower()):
                print("",end="")
            else:
                j+=1
        else:
            j+=1
    row+=1
if((i/(i+j))*100>=10 and i+j>=30):
    print("DrAiveX can EXIST in NSEC!!!")
else:
    print("DrAiveX CAN'T EXIST in NSEC!!!")
if((i/(i+1)*100)<10):
    print("Due to deficit of CE and ME guys",i,j)
        
