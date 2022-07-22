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
g3 = 1.2
ws = 0.86
wa = 0.5
cs = 1.93
ca = 1.8
inc = 0.2
i = g3 + ws*0.5

sheet1 = wb.add_sheet("Class70rw&A_3L(Left)")
sheet1.write_merge(0,0,0,1,"CLASS 70 RW",style)
sheet1.write_merge(0,0,2,3,"CLASS A",style)

for z in [0,2]:
    sheet1.write(1,z,"WHEEL 1",style)
    sheet1.write(1,z+1,"WHEEL 2",style)

a = 2
while i <= l - g1 - g2 - wa - ws*0.5 - ca:
    if i <= 7.25 - g2:
        j = 7.25 + wa*0.5
    else:
        j = i + ws*0.5 + g2 + wa*0.5 + cs
    while j <= l - g1 - ca - wa*0.5:
        w1 = i
        w2 = w1 + cs   
        w3 = j
        w4 = w3 + ca
        print(w1,w2,w3,w4)
        sheet1.write(a,0,w1,style)
        sheet1.write(a,1,w2,style)
        sheet1.write(a,2,w3,style)
        sheet1.write(a,3,w4,style)
        a += 1
        j += inc
    i += inc  
wb.save("D:\IMPROVISATION\Excel Files\Exportxl_70r&CA_3L(Left).xls")