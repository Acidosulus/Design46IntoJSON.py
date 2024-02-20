import sys
from openpyxl import load_workbook

print(sys.argv[1])


workbook = load_workbook(sys.argv[1], data_only =True)

sheet = workbook['Титульный']
data = []
for row in sheet.iter_rows(min_row=5, max_row=71, min_col=11, max_col=11):
	str = ''
	row_data = []
	for cell in row:
		if cell.value != None:
			row_data.append(cell.value)
	if len(row_data)>0:
		data.append(row_data[0])

print('\n'.join(data))

print('=================================================================')

sheet = workbook['Раздел_I__А']
data = []
for row in range(18,94+1):
	st = ''
	for col in range(105,198+1):
		readed_value = sheet.cell(row=row, column=col).value
		if readed_value != None:
			st += readed_value
	if len(st)>10:
		data.append(st)
print(data)
	



