import xlwings as xw

# Open the data workbook (use the full path if necessary)
data_wb = xw.Book('/Users/riyagaur/Downloads/Excel workbooks/Cohort Analysis.xlsx')

# List the sheet names to verify the expected sheet exists
data_sheet_names = [sheet.name for sheet in data_wb.sheets]

# Adjust the sheet name below to match exactly what appears in Excel
data_sheet = data_wb.sheets['Cohort Analysis ']

# Open the retention analysis workbook (use full path if necessary)
retention_wb = xw.Book('/Users/riyagaur/Downloads/Excel workbooks/2024.10 RadarFirst - Retention Analysis_v02.xlsx')
dest_sheet = retention_wb.sheets['Cohort Analysis ']

# Get the source range (its address tells you the exact cells that are used)
source_range = data_sheet.used_range
address = source_range.address  # e.g. "$B$2:$G$20"
print("Source range address:", address)

# Paste the data into the destination sheet at the same cell addresses
dest_sheet.range(address).value = source_range.value

# Save the retention analysis workbook
retention_wb.save()

# Attempt to close the workbooks, but catch errors if they occur
try:
    data_wb.close()
except Exception as e:
    print("Error closing data workbook:", e)

try:
    retention_wb.close()
except Exception as e:
    print("Error closing retention workbook:", e)

print("Data successfully copied from 'cohort analysis.xlsx' to the 'Cohort Analysis ' sheet in the retention analysis workbook.")
