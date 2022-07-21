import xlwt
from xlwt import Workbook
wb = Workbook()
l = 13.1
g1 = 0.15
g2 = 1.2
ww = 0.5
ca = 1.8
inc = 0.5
i=g1+ww*0.5
a = 1
sheet1 = wb.add_sheet("ClassA_3L")
sheet1.write(0,1,"WHEEL 1")
sheet1.write(0,2,"WHEEL 2")
sheet1.write(0,3,"WHEEL 3")
sheet1.write(0,4,"WHEEL 4")
sheet1.write(0,5,"WHEEL 5")
sheet1.write(0,6,"WHEEL 6")
while i<=l-g1-g2*2-ww*2-ww*0.5-ca*3+0.1:
    j=i+ww+g2+ca
    while j<=l-g1-g2-ca*2-ww*1.5+0.1:
        k=j+ww+g2+ca
        while k<=l-g1-ww*0.5-ca+0.1:
            w1 = i
            w2 = w1 + ca
            w3 = j
            w4 = w3 + ca
            w5 = k
            w6 = w5 + ca
            print(w1,w2,w3,w4,w5,w6)
            sheet1.write(a,1,w1)
            sheet1.write(a,2,w2)
            sheet1.write(a,3,w3)
            sheet1.write(a,4,w4)
            sheet1.write(a,5,w5)
            sheet1.write(a,6,w6)
            a+=1
            k+=inc
        j+=inc
    i+=inc  
wb.save("D:\IMPROVISATION\Excel Files\Exportxl.xls")