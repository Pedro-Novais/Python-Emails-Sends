import os

num = 0
try:
    for i in range (5):
        num = i
        os.mkdir(f'pdf/Line_{num}')
        print(f'Pasta Line{num} criada com sucesso')
except OSError:
    print('Erro')
