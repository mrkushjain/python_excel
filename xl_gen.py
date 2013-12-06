import xlsxwriter


workbook = xlsxwriter.Workbook('Marks.xlsx')
worksheet = workbook.add_worksheet()

result = (
    ['Ram', 96],
    ['Ravi',   99],
    ['Kush',  95],
    ['Shyam',    50],
)

row = 0
col = 0

# Iterate over the data and write it out row by row.
for name, marks in (result):
    worksheet.write(row, col,     name)
    worksheet.write(row, col + 1, marks)
    row += 1
yellow_background = workbook.add_format()

yellow_background.set_bg_color('yellow')
worksheet.write(row, 0, 'Average')
worksheet.write(row, 1, '=AVERAGE(B1:B4)',yellow_background)

workbook.close()