import xlwings as xw

# Update these paths with the actual full paths on your system.
source_path = '/Users/riyagaur/Downloads/Excel workbooks/revenue bridge.xlsx'
retention_path = '/Users/riyagaur/Downloads/Excel workbooks/2024.10 RadarFirst - Retention Analysis_v02.xlsx'

# Open the source workbook containing the "Revenue Bridge" data
source_wb = xw.Book(source_path)
# Adjust the sheet name if needed; here we assume it's "Revenue Bridge"
source_sheet = source_wb.sheets['Revenue Bridge']

# Open the retention analysis workbook
retention_wb = xw.Book(retention_path)
# Access the destination sheet; ensure the name matches exactly
dest_sheet = retention_wb.sheets['Revenue Bridge']

# Attempt to retrieve the used range from the source sheet.
# If that fails, use the contiguous range starting from A1.
try:
    source_range = source_sheet.used_range
    data = source_range.value
    address = source_range.address  # e.g., "$B$2:$G$20"
except Exception as e:
    print("Error accessing used_range, using fallback method:", e)
    data = source_sheet.range("A1").expand().value
    address = source_sheet.range("A1").expand().address

print(f"Copying data from {source_path} sheet 'Revenue Bridge' at range {address}")

# Paste the data into the destination sheet using the same address
dest_sheet.range(address).value = data

# Save the changes to the retention analysis workbook
retention_wb.save()

# Close the workbooks (wrapped in try/except to handle any closure issues)
try:
    source_wb.close()
except Exception as e:
    print("Error closing source workbook:", e)

try:
    retention_wb.close()
except Exception as e:
    print("Error closing retention workbook:", e)

print("Data successfully copied from 'revenue bridge.xlsx' to the 'Revenue Bridge' sheet in the retention analysis workbook.")
