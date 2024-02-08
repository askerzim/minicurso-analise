import pandas as pd
#instalar o openpyxl tbm

# 1-importando dados
data = pd.read_excel("tabela/VendaCarros.xlsx") #meu DataFrame original

# 2-selecionando colunas específicas do dataframe (criarei uma nova para fazer isso)
df = data[["Fabricante", "ValorVenda", "Ano"]] #selecionar apenas as tabelas que me interessam e trabalhar com ela
print(df) 

# 3-criando a tabela pivô
pivot_table = df.pivot_table(
    index="Ano", #na tabela o indice começa com:0,1,2,3,...fim. Queremos mudar o índice para o ano da planilha, então será ano
    columns="Fabricante", #nas colunas quero trabalhar com o fabricante
    values="ValorVenda", #nos valores, quero trabalhar com o valor venda
    aggfunc='sum' #por fim, quero uma função para somar os valores ao decorrer dos anos de cada fabricante
)

print(pivot_table)

# 4-exportando tabela pivô em arquivo excel

pivot_table.to_excel("tabela/pivot_table.xlsx", "Relatório") #1p: caminho para salvar. 2p: nome do arquivo