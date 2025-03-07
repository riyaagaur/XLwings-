import xlwings as xw

# Update these paths with the actual full paths on your system.
source_path = '/Users/riyagaur/Downloads/Excel workbooks/Cohort by Quarter.xlsx'
retention_path = '/Users/riyagaur/Downloads/Excel workbooks/2024.10 RadarFirst - Retention Analysis_v02.xlsx'

# Open the source workbook containing the "Cohort by Quarter" data
source_wb = xw.Book(source_path)
source_sheet = source_wb.sheets['Cohort by Quarter']

# Open the retention analysis workbook
retention_wb = xw.Book(retention_path)
dest_sheet = retention_wb.sheets['Cohort by Quarter']

# Try to retrieve the used range (to preserve the original cell positions)
try:
    source_range = source_sheet.used_range
    data = source_range.value
    address = source_range.address  # e.g., "$B$2:$G$20"
except Exception as e:
    print("Error accessing used_range, using fallback:", e)
    data = source_sheet.range("A1").expand().value
    address = source_sheet.range("A1").expand().address

print(f"Copying data from {source_path} sheet 'Cohort by Quarter' at range {address}")

# Paste the data into the destination sheet using the same address
dest_sheet.range(address).value = data

# Save the changes to the retention analysis workbook
retention_wb.save()

# Close the workbooks (wrapped in try/except in case of errors)
try:
    source_wb.close()
except Exception as e:
    print("Error closing source workbook:", e)
    
try:
    retention_wb.close()
except Exception as e:
    print("Error closing retention workbook:", e)

print("Data successfully copied from 'Cohort by Quarter.xlsx' to the 'Cohort by Quarter' sheet in the retention analysis workbook.")
