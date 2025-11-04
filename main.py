## importa a biblioteca panda e quando for usar funções dela vou me refirir a ela como pd
import pandas as pd


print("")
print("")

## de padrão não seria mostrado para mim todas as linhas apenas algumas, então eu adicionei essa parte que obriga o pandas a mostrar todas as linhas e colunas que tiver na tabela
pd.set_option('display.max_rows', None)     
pd.set_option('display.max_columns', None)  
pd.set_option('display.width', 1000)


## é aqui onde vai o arquivo para a análise 
df = pd.read_excel(r"Atividade-29.10.xlsx", 
sheet_name = "Página1",
skiprows=2,
)


## remove espaçamento para não ter chance de dar erro nos nomes
df.columns = df.columns.str.strip()

## é aqui onde eu ponho as colunas que vão ser utilizadas para leitura
colunas_desejadas = [
    f"ID", 
    f"Data da Venda", 
    f"Nome do Cliente", 
    f"Valor Total", 
    f"Forma de pagamento", 
    f"Produto vendido", 
    f"Valor pago", 
    f"Valor de troco", 
    f"Vendedor"
]

df = df[colunas_desejadas]


## printa a planilha com todas as colunas e linha
print(df)


print("")
print("")

## aqui tira os R$ que iria causar erros quando fosse fazer a análise da coluna que contém os valores
df['Valor Total'] = df['Valor Total'].astype(str).str.replace('R$', '', regex=False).str.replace(',', '.', regex=False).str.strip()
## caso de algum erro na leitura de algum valor, sera mostrado um NaN
df['Valor Total'] = pd.to_numeric(df['Valor Total'], errors='coerce')

print("\n=======================================================")
print("                 O produto mais vendido                  ")
print("=======================================================")

## usando a coluna valor total ele pega o produto mais vendido
receita_por_produto = df.groupby('Produto vendido')['Valor Total'].sum().sort_values(ascending=False)

print(receita_por_produto.to_string())


print("")
print("")

print("\n=======================================================")
print("           Maior valor vendido por vendedor"             )
print("=======================================================")

## usando a coluna valor total ele pega o vendedor com mais vendas
receita_por_vendedor = df.groupby('Vendedor')['Valor Total'].sum().sort_values(ascending=False)

print(receita_por_vendedor.to_string())


print("")
print("")


## aqui tira os R$ que iria causar erros quando fosse fazer a análise da coluna que contém os valores
df['Valor pago'] = df['Valor pago'].astype(str).str.replace('R$', '', regex=False).str.replace(',', '.', regex=False).str.strip()
## caso de algum erro na leitura de algum valor, sera mostrado um NaN
df['Valor pago'] = pd.to_numeric(df['Valor pago'], errors='coerce')

maiorValorPagoPorCliente = df.groupby('Nome do Cliente')['Valor pago'].sum().sort_values(ascending=False)

print("\n=======================================================")
print("                 Qual cliente pagou mais  "              )
print("=======================================================")
print(maiorValorPagoPorCliente)

menorValorPagoPorCliente = df.groupby('Nome do Cliente')['Valor pago'].sum().sort_values(ascending=True)
print("\n=======================================================")
print("                 Qual cliente pagou menos  "             )
print("=======================================================")
print(menorValorPagoPorCliente)


maiorTroco = df['Valor de troco'].max()
print("\n=======================================================")
print("    Qual o troco mais alto fornecido por um cliente  ")
print("=======================================================")
print(maiorTroco)