import openpyxl
from string import Template
import codecs 
import webbrowser

wb = openpyxl.load_workbook('teste_faturamento.xlsx')

def search(title, values):
    for atrb in values:
        print("")
        print(title)
        print(atrb)

aba = wb['TESTE']
valor_title = []
valor = []
i = 0
for linha in aba:
    if(i>0):
        valor.append([])
    for celula in linha:
        posC = celula.column
        posR = celula.row
        valorT = celula.value
        if(posR < 2):
            valor_title.append(valorT)
            i = 1
        if(posR > 1):
            valor[posR-2].append(valorT)
            print()
            #print(valor_title)
            #print(valor[posR-2])   

if __name__ == '__main__':
    search(valor_title, valor)

#print("")
#print(valor[1])

file = codecs.open("index.html", 'r', "utf-8") 
print(file.read()) 

webbrowser.open('index.html')  