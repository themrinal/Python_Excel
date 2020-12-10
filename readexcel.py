import openpyxl

wb=openpyxl.load_workbook("wrksht.xlsx")
#this program will return max_column and max_row in your worksheet
print(wb.active.title)
sheet=wb['First']
rows=sheet.max_row
cols=sheet.max_column
# print(rows,cols)

for i in range(1,rows+1):       #this will goes upto last rows
    for j in range(1,cols+1):   #this will goes upto last cols
        print(sheet.cell(i,j).value)