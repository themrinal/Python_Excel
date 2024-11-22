import openpyxl

# Load the Excel workbook into the 'wrk' variable
wrk = openpyxl.load_workbook("wrksht.xlsx")  

# Print the type of the loaded workbook to confirm successful loading
# This will output the class associated with the 'wrk' variable
print(type(wrk))  

# Retrieve and print a list of all worksheet names in the workbook
sheets = wrk.sheetnames
print(sheets)  # This will return a list of all the sheet names in the Excel file

# Print the name of the active sheet in the workbook
# The active sheet is the one that is currently selected by default when the file opens
print(wrk.active.title)  # Example: 'Sheet1'


