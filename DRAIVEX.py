import xlrd
loc = "E:/Important/Project/Draivex/test data/dataset.xlsx"
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
i=0
j=0
row=1
flag=0
temp=''
while(row<sheet.nrows):
    flag=0
    if ((sheet.cell_value(row, 4)=="ME") or(sheet.cell_value(row, 4)=="CE")):        
        if sheet.cell_value(row, 3)==("3rd" or "4th"):
            x=sheet.cell_value(row, 2)
            y=sheet.cell_value(row,1)
            name=y.split(" ")
            for k in range(len(name)):
                temp=name[k]
                if temp.lower() in x.lower():
                    flag=1
                    break                
            if (flag==1):
                print("",end="")
            else:
                i+=1
        else:
            i+=1
    else:
        if sheet.cell_value(row, 3)==("3rd" or "4th"):
            x=sheet.cell_value(row, 2)
            y=sheet.cell_value(row,1)
            name=y.split(" ")
            for k in range(len(name)):
                temp=name[k]
                if temp.lower() in x.lower():
                    flag=1
                    break      
            if (flag==1):
                print("",end="")
            else:
                j+=1
        else:
            j+=1
    row+=1
print("CE and ME students combined = ",i)
print("other students = ",j)
if((i/(i+j))*100>=10 and i+j>=30):
    print("DrAiveX can EXIST in NSEC!!!")
else:
    print("DrAiveX CAN'T EXIST in NSEC!!!")
if((i/(i+j)*100)<10):
    print("Due to deficit of CE and ME students")
        
