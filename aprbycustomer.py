import xlwings as xw

# Update these paths with your actual file locations.
source_path = '/Users/riyagaur/Downloads/Excel workbooks/aprbycustomer.xlsx'
retention_path = '/Users/riyagaur/Downloads/Excel workbooks/2024.10 RadarFirst - Retention Analysis_v02.xlsx'

# Open the source workbook containing the "APR By Customer" data
source_wb = xw.Book(source_path)
# Adjust the source sheet name if needed; here we assume it's "APR By Customer"
source_sheet = source_wb.sheets['ARR By Customer']

# Open the retention analysis workbook
retention_wb = xw.Book(retention_path)
# Access the destination sheet; ensure the name matches exactly, here it's "ARR by Customer"
dest_sheet = retention_wb.sheets['ARR by Customer']

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

print(f"Copying data from {source_path} sheet 'APR By Customer' at range {address}")

# Paste the data into the destination sheet using the same address to preserve the layout
dest_sheet.range(address).value = data

# Save the changes to the retention analysis workbook
retention_wb.save()

# Close the workbooks (using try/except to catch any closure issues)
try:
    source_wb.close()
except Exception as e:
    print("Error closing source workbook:", e)

try:
    retention_wb.close()
except Exception as e:
    print("Error closing retention workbook:", e)

print("Data successfully copied from 'aprbycustomer.xlsx' to the 'ARR by Customer' sheet in the retention analysis workbook.")
