import openpyxl

arq = openpyxl.load_workbook('dados.xlsx')

plan = arq['Planilha1']

x = plan['A3'].value

print(x)