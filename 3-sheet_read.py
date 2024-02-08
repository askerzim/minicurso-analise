from openpyxl import load_workbook

# 1- Lê pasta de trabalho e planilha
wb = load_workbook("tabela/pivot_table.xlsx") #carrego o workbook e atribuo a var wb
sheet = wb["Relatório"] #passo o nome da planilha que quero trabalhar 

# 2- Acessando um valor específico através da tabela
#print(sheet["B3"].value)
#a partir disso, conseguimos acessar qualquer valor de cedula com as coordenadas

# 3-Iterando valores por meio de loop
for i in range(2,6):
    ano = sheet["A%s" %i].value #o %s sera o valor da var de controle i, oq ira gerar algo dinamico na varredura de valores
    am = sheet["B%s" %i].value
    bt = sheet["C%s" %i].value
    print(f'No ano de {ano} o Aston Martin vendeu {am}, enquanto o Bentley vendey {bt}.')