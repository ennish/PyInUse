import xlrd
import xlwt

wb = xlrd.open_workbook(r'1.xlsx')
sheet1 = wb.sheet_by_index(0)

totalRows = sheet1.nrows
print('totalRows:')
print(totalRows)
#表头，结尾
heads = sheet1.row_values(0,0,6)
tails = sheet1.row_values(totalRows-1,0,6)

print(heads)
print(tails)

left_top50,right_top50,left_extra,right_extra = [],[],[],[]
print('reading\n')
#读取初始sheet i行号
for i in range (1,51):
    left_top50.append(sheet1.row_values(i,0,6))

print("length of left top:")
print (len(left_top50))

for i in range (51,101):
    right_top50.append(sheet1.row_values(i,0,6))
print("length of right top:")
print (len(right_top50))

for i in range (101,totalRows-2):
    if ((i-101)//54)%2 == 0:
        left_extra.append(sheet1.row_values(i,0,6))
    else:
        right_extra.append(sheet1.row_values(i,0,6))

# print(left_extra)
# print(right_extra)

print("length of left extra:")
print (len(left_extra))

print ("length of right extra:")
print (len(right_extra))

wb2 = xlwt.Workbook()
sheet2 = wb2.add_sheet(u'sheet2',cell_overwrite_ok=True)

sheet2.col(2).width=256*16
sheet2.col(8).width=256*16

num_style = xlwt.easyxf(num_format_str='#')
print('\n writing \n')
#count: 新Excel行数(包括表头，表尾),i:新Excel数据表行数
count = 0
for i in range(0,len(left_extra)+ 50 ):
    if i == 0:
        for j in range(0,6):
            sheet2.write(count ,j ,heads[j])
        for j in range(6,12):
            sheet2.write(count ,j,heads[j%6])
        count += 1
         
    if i < 50:
        for j in range(0,6):
            if j==2:
                sheet2.write(count,j,left_top50[i][j],num_style)
            else:
                sheet2.write(count,j,left_top50[i][j])
        for j in range(6,12):
            if j==8:
                sheet2.write(count,j,right_top50[i][j%6],num_style)
            else:
                sheet2.write(count,j,right_top50[i][j%6])
        count += 1
        
    if i >= 50:
        # print("count:"+str(count))
        # print("\n")
        # print("i:"+str(i))
        # print("\n")

        if i < len(left_extra) + 50:
            if (i-50)%54 == 0:
                for j in range(0,6):
                    sheet2.write(count ,j ,heads[j])
                if i < len(right_extra)+50:
                    for j in range(6,12):
                        sheet2.write(count ,j ,heads[j%6])
                count += 1

            for j in range(0,6):
                if j==2:
                    sheet2.write(count,j,left_extra[i-50][j],num_style)
                else:
                    sheet2.write(count,j,left_extra[i-50][j])
            if i < len(right_extra) + 50:
                for j in range(6,12):
                    if j==8:
                        sheet2.write(count,j,right_extra[i-50][j%6],num_style)
                    else:
                        sheet2.write(count,j,right_extra[i-50][j%6])
            #print tail
            if  i == len(right_extra) + 50:
                for j in range(6,12):
                    sheet2.write(count ,j ,tails[j%6])
            count += 1

wb2.save('1_export.xlsx')

#算法设计时，先要想清楚，明晰具体规则
