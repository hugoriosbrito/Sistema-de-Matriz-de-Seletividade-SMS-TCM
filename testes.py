import openpyxl as xl
import xlwings as xw

wb = xw.Book('dados\\Matriz modelo - REV3 - Otimizado.xlsx')
sheet = wb.sheets['SÍNTESE MUN.']

lista = ['SIM','NÃO','NÃO', 'SIM']

sheet.range('G11:G14').value = lista

try:
    wb.save('dados\\Matriz modelo - REV3 - Otimizado.xlsx')
    print(sheet['G11:G14'])
    print('salvo com sucesso')
except Exception as e:
    print(f'Erro: {e}')

wb.close()