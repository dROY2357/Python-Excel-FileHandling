import xlrd
import xlwt
from xlwt import Workbook

loc = "COPO - Data April 2019 to April 2020 - CSE.xlsx"
wb = xlrd.open_workbook(loc)
output_wb = Workbook()

sheet1 = wb.sheet_by_index(0)
sheet2 = wb.sheet_by_index(1)
sheet3 = wb.sheet_by_index(2)

out_sheet1 = output_wb.add_sheet('APR19')
out_sheet2 = output_wb.add_sheet('OCT19')
out_sheet3 = output_wb.add_sheet('APR20')

sub_name = str(input("Enter the subject name you want to extract data of:"))

style_header = xlwt.easyxf('font: bold 1; align: wrap on, vert centre, horiz center')
style_cell = xlwt.easyxf('align: wrap on, vert center, horiz center')
sheet_cell_nowrap = xlwt.easyxf('align: vert center, horiz center')

column_header_list = ['Name', 'SRN', 'IA_Marks_q1', 'IA_Marks_q2', 'IA_Marks_q3', 'IA_Marks_q4', 'SEE_Marks']

def extractData(sheet_no, output_sheet_no):

    for i in range(7):
        output_sheet_no.write(0, i, column_header_list[i], style_header)

    name_list = []
    srn_list = []
    ia_marks_list = []
    see_marks_list = []

    for i in range(sheet_no.nrows):
        if (str(sheet_no.cell_value(i, 11)) == str(sub_name)) and ("R16CS001" <= str(sheet_no.cell_value(i, 8)) <= "R16CS600" or str(sheet_no.cell_value(i, 8)) >= "R17CS800"):
            name_list.append(sheet_no.cell_value(i, 9))
            srn_list.append(sheet_no.cell_value(i, 8))
            ia_marks_list.append(sheet_no.cell_value(i, 14))
            see_marks_list.append(sheet_no.cell_value(i, 13))

    for i in range(len(name_list)):
        output_sheet_no.write(i+1, 0, name_list[i], sheet_cell_nowrap)
        output_sheet_no.write(i+1, 1, srn_list[i], sheet_cell_nowrap)
        output_sheet_no.write(i+1, 2, (float(ia_marks_list[i])/2.0), style_cell)
        output_sheet_no.write(i+1, 4, (float(ia_marks_list[i])/2.0), style_cell)
        output_sheet_no.write(i+1, 6, float(see_marks_list[i]), style_cell)

extractData(sheet1, out_sheet1)
extractData(sheet2, out_sheet2)
extractData(sheet3, out_sheet3)

out_loc = str(input("Enter the output file name with .xls extention, that you want to store output in:"))

output_wb.save(out_loc)
