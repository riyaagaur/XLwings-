import xlwings as xw

# Open the data workbook (the file with your source data)
data_wb = xw.Book('OG Source.xlsx')
# Replace 'og source' with the exact sheet name if it differs
data_sheet = data_wb.sheets['OG Source']

# Open the retention analysis workbook
retention_wb = xw.Book('2024.10 RadarFirst - Retention Analysis_v02.xlsx')
# Make sure the destination sheet name matches exactly; adjust if needed
dest_sheet = retention_wb.sheets['OG Source']

# Copy all data from the data workbook's sheet using the used range
data = data_sheet.used_range.value

# Paste the data into the destination sheet starting at cell A1
dest_sheet.range('A1').value = data

# Save the retention analysis workbook (this overwrites the existing file)
retention_wb.save()

# Optionally close both workbooks
data_wb.close()
retention_wb.close()

print("Data successfully copied from OG Source.xlsx to the 'og source' sheet in the retention analysis workbook.")
