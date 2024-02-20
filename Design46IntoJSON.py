import sys
from openpyxl import load_workbook

print(sys.argv[1])


print('===============================================================')
pattern_content = open(file='pattern.json', mode='r', encoding='utf-8').read()
print(pattern_content)
print('===============================================================')

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

#print('\n'.join(data))
pattern_content = pattern_content.replace('{{ Титульный }}','\n'.join(data))

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
#print('\n'.join(data))
pattern_content = pattern_content.replace('{{ Раздел_I__А }}','\n'.join(data))


print('=================================================================')

sheet = workbook['Раздел_I__Б']
data = []
for row in range(18,94+1):
	st = ''
	for col in range(19, 19+1):
		readed_value = sheet.cell(row=row, column=col).value
		if readed_value != None:
			st += readed_value
	if len(st)>10:
		data.append(st)
#print('\n'.join(data))
pattern_content = pattern_content.replace('{{ Раздел_I__Б }}','\n'.join(data))

#print('=================================================================')

sheet = workbook['Раздел_I__В']
data = []
for row in range(18,61+1):
	st = ''
	for col in range(28, 44+1):
		readed_value = sheet.cell(row=row, column=col).value
		if readed_value != None:
			st += readed_value
	if len(st)>10:
		data.append(st)
#print('\n'.join(data))
pattern_content = pattern_content.replace('{{ Раздел_I__В }}','\n'.join(data))

#print('=================================================================')

sheet = workbook['Раздел_II__А_(ТИС)']
data = []
for row in range(18,39+1):
	st = ''
	for col in range(67, 123):
		readed_value = sheet.cell(row=row, column=col).value
		if readed_value != None:
			st += readed_value
	if len(st)>10:
		data.append(st)
#print('\n'.join(data))
pattern_content = pattern_content.replace('{{ Раздел_II__А_(ТИС) }}','\n'.join(data))

#print('=================================================================')

sheet = workbook['Раздел_II__Б_(ТИС)']
data = []
for row in range(18,39+1):
	st = ''
	for col in range(15, 15+1):
		readed_value = sheet.cell(row=row, column=col).value
		if readed_value != None:
			st += readed_value
	if len(st)>10:
		data.append(st)
#print('\n'.join(data))
pattern_content = pattern_content.replace('{{ Раздел_II__Б_(ТИС) }}','\n'.join(data))

#print('=================================================================')

sheet = workbook['Раздел_III']
data = []
for row in range(18,45+1):
	st = ''
	for col in range(12, 12+1):
		readed_value = sheet.cell(row=row, column=col).value
		if readed_value != None:
			st += readed_value
	if len(st)>10:
		data.append(st)
#print('\n'.join(data))
pattern_content = pattern_content.replace('{{ Раздел_III }}','\n'.join(data))

#print('=================================================================')

sheet = workbook['Раздел_IV']
data = []
for row in range(18,41+1):
	st = ''
	for col in range(12, 12+1):
		readed_value = sheet.cell(row=row, column=col).value
		if readed_value != None:
			st += readed_value
	if len(st)>10:
		data.append(st)
#print('\n'.join(data))
pattern_content = pattern_content.replace('{{ Раздел_IV }}','\n'.join(data))

open(file=sys.argv[1]+'.json', mode='a', encoding='utf-8').write(pattern_content)

input("Press Enter to continue...")


