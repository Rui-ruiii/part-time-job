import xlrd
import random
#读取文件

data = xlrd.open_workbook(r'excel-forlist.xlsx')

read_sheet1 = data.sheets()[0]
read_sheet2 = data.sheets()[1]
read_sheet3 = data.sheets()[2]
read_sheet4 = data.sheets()[3]
read_sheet5 = data.sheets()[4]

lista = []
listb = []
listc = []
listd = []
liste = []


a = int(input("enter the rows of sheet1:"))
b = int(input("enter the rows of sheet2:"))
c = int(input("enter the rows of sheet3:"))
d = int(input("enter the rows of sheet4:"))
e = int(input("enter the rows of sheet5:"))

def Start():
    for rows in range(0,a):
        lista.append(read_sheet1.cell(rows,0).value)
        num1 = random.randint(0,a)
    print(lista[num1])

def Mid1():
    for rows in range(0,b):
        listb.append(read_sheet2.cell(rows,0).value)
        num2 = random.randint(0,b)
    print(listb[num2])

def Mid2():
    for rows in range(0,c):
        listc.append(read_sheet3.cell(rows,0).value)
        num3 = random.randint(0,c)
    print(listc[num3])

def Mid3():
    for rows in range(0,d):
        listd.append(read_sheet4.cell(rows,0).value)
        num4 = random.randint(0,d)
    print(listd[num4])

def End():
    for rows in range(0,e):
        liste.append(read_sheet5.cell(rows,0).value)
        num5 = random.randint(0,e)
    print(liste[num5])

Start()
Start()
Start()
Start()
Start()
Start()
print("\n\t")
Mid1()
Mid1()
Mid1()
Mid1()
Mid1()
Mid1()
print("\n\t")
Mid2()
Mid2()
Mid2()
Mid2()
Mid2()
Mid2()
print("\n\t")
Mid3()
Mid3()
Mid3()
Mid3()
Mid3()
Mid3()
print("\n\t")
End()
End()
End()
End()
End()
End()