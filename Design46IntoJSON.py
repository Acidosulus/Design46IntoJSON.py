import sys
from openpyxl import load_workbook

print(sys.argv[1])

file_source = sys.argv[1]
if file_source[-3:].upper() == 'ODS':

	print('!!!!!!!!!!!!!!!!!!!!!!!!!')
	import win32com.client as win32


	# Сохранение файла в другом формате с помощью win32com
	excel_app = win32.Dispatch('Excel.Application')
	wb_xl = excel_app.Workbooks.Open(file_source)
	file_source += '.xlsx'
	wb_xl.SaveAs(file_source, FileFormat=51)
	wb_xl.Close()
	excel_app.Quit()


print('===============================================================')
pattern_content = open(file='pattern.json', mode='r', encoding='utf-8').read()
#print(pattern_content)
print('===============================================================')

workbook = load_workbook(file_source, data_only =True)

print('=================================================================')

sheet = workbook['Титульный']
data = []
for row in range(5,71+1):
	st = ''
	for col in range(11,11+1):
		readed_value = sheet.cell(row=row, column=col).value
		if readed_value != None:
			st += readed_value
	if len(st)>10:
		data.append(st)
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
	for col in range(67, 123+1):
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

open(file=file_source+'.json', mode='a', encoding='utf-8').write(pattern_content)

#input("Press Enter to continue...")


