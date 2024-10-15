import xlsxwriter

# Create a workbook and add a worksheet
workbook = xlsxwriter.Workbook('sales_report.xlsx')
worksheet = workbook.add_worksheet()

# Some data to write
data = [
    ['Date', 'Product', 'Quantity', 'Price'],
    ['2024-10-01', 'Product A', 5, 10.50],
    ['2024-10-02', 'Product B', 3, 15.75],
    ['2024-10-03', 'Product C', 8, 12.00],
    ['2024-10-03', 'Product D', 8, 12.00],
    ['2024-10-03', 'Product E', 8, 12.00]
]

# Write the data to the worksheet
for row_num, row_data in enumerate(data):
    for col_num, cell_data in enumerate(row_data):
        worksheet.write(row_num, col_num, cell_data)

# Close the workbook
workbook.close()
