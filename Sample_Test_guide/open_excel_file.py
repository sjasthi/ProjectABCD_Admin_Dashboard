import openpyxl

# Open the Excel file
file_path = 'Project ABCD dresses.xlsx'
workbook = openpyxl.load_workbook(file_path)

# Select a specific sheet from the Excel file
sheet = workbook['sheet1'] # Replace with sheet name

# Read data from the Excel sheet
for row in sheet.iter_rows(values_only=True):
    for cell in row:
        print(cell)
