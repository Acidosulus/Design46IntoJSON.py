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



sheet = workbook['Раздел_I__А']
data = []
for row in sheet.iter_rows(min_row=18, max_row=94, min_col=105, max_col=198):
	str = ''
	row_data = []
	for cell in row:
		if cell.value != None:
			row_data.append(cell.value)
	if len(row_data)>0:
		data.append(row_data[0])

print(data)



