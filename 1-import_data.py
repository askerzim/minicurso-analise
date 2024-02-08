import pandas as pd
#instalar o openpyxl tbm

# 1-importando dados
data = pd.read_excel("tabela/VendaCarros.xlsx")
print(data)

# 2-listar os 5 (padrão se nao passar parametros) primeiros registros
print(data.head())

# 3-listar os 5 (padrão se nao passar parametros) últimos registros
print(data.tail())

# 4-contagem de valores por Fabricante (acessar uma coluna e contabilizar)
print(data['Fabricante'].value_counts())

