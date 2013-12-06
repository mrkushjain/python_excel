import xlsxwriter


workbook = xlsxwriter.Workbook('Marks.xlsx')
worksheet = workbook.add_worksheet()

result = (
    ['Ram', 96],
    ['Ravi',   2],
    ['Kush',  5],
    ['Shyam',    5],
)

row = 0
col = 0
sum = 0

# Iterate over the data and write it out row by row.
for name, marks in (result):
    worksheet.write(row, col,     name)
    worksheet.write(row, col + 1, marks)
    sum += marks
    row += 1

no_of_students = len(result)
average = float(sum)/no_of_students

background = workbook.add_format()
if average > 50:
	background.set_bg_color('yellow')
else:
	background.set_bg_color('red')
worksheet.write(row, 0, 'Average')
worksheet.write(row, 1, average,background)

workbook.close()