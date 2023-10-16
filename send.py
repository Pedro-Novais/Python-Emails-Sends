import openpyxl;
wb = openpyxl.Workbook()
aba = wb['Sheet']
aba.title = 'Usuários'
aba['A1'] = 'Nome'
aba['A2'] = 'João'
aba['A3'] = 'José'
aba['A4'] = 'Pedro'
aba['B1'] = 'Idade'
aba['B2'] = 25
aba['B3'] = 28
aba['B4'] = 19
aba['C1'] = 'Email'
aba['C2'] = "teste@gmail.com"
wb.save('Dados.xlsx')