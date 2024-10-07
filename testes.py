import xlwings as xw

wb = xw.Book('Matriz modelo - REV3.xlsx')
sheet = wb.sheets['SÍNTESE MUN.']

valor1 = 'NÃO'
valor2 = 'NÃO'

valores = sheet['G14'].value = 'SIM'
wb.save()
print(valores)

wb.close()