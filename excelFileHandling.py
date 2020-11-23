import xlrd
import xlwt
from xlwt import Workbook

loc = "COPO - Data April 2019 to April 2020 - CSE.xlsx"
wb = xlrd.open_workbook(loc)
output_wb = Workbook()

sheet1 = wb.sheet_by_index(0)
sheet2 = wb.sheet_by_index(1)
sheet3 = wb.sheet_by_index(2)

sub_name = str(input("Enter the subject name you want to extract data of:"))

style = xlwt.easyxf('font: bold 1')

column_header_list = ['Name', 'SRN', 'IA_Marks', 'SEE_Marks']

out_sheet1 = output_wb.add_sheet('APR19')

for i in range(0,4):
    out_sheet1.write(0, i, column_header_list[i], style)

name_list = []
srn_list = []
ia_marks_list = []
see_marks_list = []

for i in range(sheet1.nrows):
    if (str(sheet1.cell_value(i, 11)) == str(sub_name)) and ("R16CS001" <= str(sheet1.cell_value(i, 8)) <= "R16CS600" or str(sheet1.cell_value(i, 8)) >= "R17CS800"):
        name_list.append(sheet1.cell_value(i, 9))
        srn_list.append(sheet1.cell_value(i, 8))
        ia_marks_list.append(sheet1.cell_value(i, 14))
        see_marks_list.append(sheet1.cell_value(i, 13))

for i in range(len(name_list)):
    out_sheet1.write(i+1, 0, name_list[i])
    out_sheet1.write(i+1, 1, srn_list[i])
    out_sheet1.write(i+1, 2, ia_marks_list[i])
    out_sheet1.write(i+1, 3, see_marks_list[i])


output_wb.save("Out.xls")
