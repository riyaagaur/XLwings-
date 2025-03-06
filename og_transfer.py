import xlwings as xw

# Open the retention workbook (adjust the file path if needed)
source_wb = xw.Book('2024.10 RadarFirst - Retention Analysis_v02.xlsx')
source_sheet = source_wb.sheets['Og Source']

# Create a new workbook which will serve as the copy
dest_wb = xw.Book()
dest_sheet = dest_wb.sheets[0]  # Get the first (default) sheet

# Rename the default sheet to "og source"
dest_sheet.name = 'og source'

# Copy data from the source sheet (using the used range; adjust range if needed)
data = source_sheet.used_range.value

# Paste the data into the destination sheet starting at cell A1
dest_sheet.range("A1").value = data

# Save the new workbook to a file
dest_wb.save('Og_Source_Copy.xlsx')

# Optionally, close the workbooks
source_wb.close()
dest_wb.close()

print("Data successfully copied to 'Og_Source_Copy.xlsx'!")
