import win32com.client as win32
import pandas as pd

# Define the date range for filter
start_date = '01-01-2023'
end_date = '30-12-2023'

# Open the Excel workbook
excel = win32.gencache.EnsureDispatch('Excel.Application')
workbook = excel.Workbooks.Open(r'https://my.sharepoint.com/yourfile.xlsm')

# Filter the dataframe by the desired date range
workbook = workbook.loc[(workbook['date_column'] >= start_date) & (workbook['date_column'] <= end_date)]

# Check if Summary sheet already exist, if not create one
if "Summary" not in [sheet.Name for sheet in workbook.Sheets]:
    summary_sheet = workbook.Sheets.Add(After=workbook.Sheets(workbook.Sheets.Count))
    summary_sheet.Name = "Summary"
else:
    summary_sheet = workbook.Sheets("Summary")
    summary_sheet.UsedRange.Clear()

# Loop through all sheets in the workbook (excluding "Summary" sheet)
for sheet in workbook.Sheets:
    if sheet.Name != "Summary":
        # Copy headers
        sheet.Range(sheet.Rows(1), sheet.Rows(1)).Copy()
        summary_sheet.Range(summary_sheet.Rows(summary_sheet.Rows.Count), summary_sheet.Rows(summary_sheet.Rows.Count)).Offset(1, 0).PasteSpecial()
        # Get last row of data for today's date
        last_row = sheet.Cells(sheet.Rows.Count, 1).End(win32.constants.xlUp).Row
        # Loop through each row in the current sheet
        for row in range(2, last_row + 1):
            # Check if the date in the first column matches today's date
            if sheet.Cells(row, 1).Value == date_filter:
                # Copy the current row and paste it into the "Summary" sheet
                sheet.Range(sheet.Rows(row), sheet.Rows(row)).Copy()
                summary_sheet.Range(summary_sheet.Rows(summary_sheet.Rows.Count), summary_sheet.Rows(summary_sheet.Rows.Count)).Offset(1, 0).PasteSpecial()
                # Set the number format for the date column
                summary_sheet.Cells(summary_sheet.Rows.Count, 1).End(win32.constants.xlUp).NumberFormat = "dd/mm/yyyy"

# Autofit the columns
summary_sheet.Columns.AutoFit()

# Save the workbook
workbook.Save()

# Close Excel
excel.Application.Quit()
