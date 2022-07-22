import xlwt
from xlwt import Workbook
wb = Workbook()
alignment = xlwt.Alignment()
alignment.horz = xlwt.Alignment.HORZ_CENTER
alignment.vert = xlwt.Alignment.VERT_CENTER
style = xlwt.XFStyle()
style.alignment = alignment
l = 13.1
g1 = 0.15
g2 = 1.2
wa = 0.5
ca = 1.8
inc = 0.2
i = g1 + wa*0.5

sheet1 = wb.add_sheet("ClassA_3L")
sheet1.write_merge(0,0,0,1,"CLASS A",style)
sheet1.write_merge(0,0,2,3,"CLASS A",style)
sheet1.write_merge(0,0,4,5,"CLASS A",style)

for z in [0,2,4]:
    sheet1.write(1,z,"WHEEL 1",style)
    sheet1.write(1,z+1,"WHEEL 2",style)

a = 2
while i <= l - g1 - g2*2 - wa*2 - wa*0.5 - ca*3:
    j = i + wa + g2 + ca
    while j <= l - g1 - g2 - ca*2 - wa*1.5:
        k = j + wa + g2 + ca
        while k <= l - g1 - wa*0.5 - ca:
            w1 = i
            w2 = w1 + ca
            w3 = j
            w4 = w3 + ca
            w5 = k
            w6 = w5 + ca
            print(w1,w2,w3,w4,w5,w6)
            sheet1.write(a,0,w1,style)
            sheet1.write(a,1,w2,style)
            sheet1.write(a,2,w3,style)
            sheet1.write(a,3,w4,style)
            sheet1.write(a,4,w5,style)
            sheet1.write(a,5,w6,style)
            a += 1
            k += inc
        j += inc
    i += inc  
wb.save("D:\IMPROVISATION\Excel Files\Exportxl_CA_3L.xls")