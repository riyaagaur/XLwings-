import xlwings as xw

# Update these paths with your actual file locations.
source_path = '/Users/riyagaur/Downloads/Excel workbooks/control panel.xlsx'
retention_path = '/Users/riyagaur/Downloads/Excel workbooks/2024.10 RadarFirst - Retention Analysis_v02.xlsx'

# Open the source workbook
source_wb = xw.Book(source_path)
source_sheet_names = [sheet.name for sheet in source_wb.sheets]
print("Source workbook sheets:", source_sheet_names)

# Use the expected sheet name; update if necessary
source_sheet_name = 'Control Panel'
if source_sheet_name not in source_sheet_names:
    raise ValueError(f"Sheet '{source_sheet_name}' not found in {source_path}.")
source_sheet = source_wb.sheets[source_sheet_name]

# Open the retention analysis workbook
retention_wb = xw.Book(retention_path)
dest_sheet = retention_wb.sheets['Control Panel']

# Try to retrieve the used range; if that fails, fall back to expanding from A1
try:
    source_range = source_sheet.used_range
    data = source_range.value
    address = source_range.address
except Exception as e:
    print("Error accessing used_range, using fallback method:", e)
    data = source_sheet.range("A1").expand().value
    address = source_sheet.range("A1").expand().address

print(f"Copying data from {source_path} sheet '{source_sheet_name}' at range {address}")
dest_sheet.range(address).value = data

# Save the retention analysis workbook
retention_wb.save()

# Close the workbooks with error handling
try:
    source_wb.close()
except Exception as e:
    print("Error closing source workbook:", e)

try:
    retention_wb.close()
except Exception as e:
    print("Error closing retention workbook:", e)

print("Data successfully copied from 'control panel.xlsx' to the 'Control Panel' sheet in the retention analysis workbook.")
