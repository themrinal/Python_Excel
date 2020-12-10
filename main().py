import openpyxl

wrk=openpyxl.load_workbook("wrksht.xlsx") #this method will load your worksheet into a var
print(type(wrk)) #basically it will print a class that is assigned with wrk var

sheets=wrk.sheetnames
print(sheets) # this will return how many sheets in that file. (it will print lists)

#First we have to see which sheet are active.
print(wrk.active.title) #sheet1

