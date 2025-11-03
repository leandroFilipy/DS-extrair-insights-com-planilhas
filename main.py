import pandas as pd


print("")
print("")


pd.set_option('display.max_rows', None)     
pd.set_option('display.max_columns', None)  
pd.set_option('display.width', 1000)

df = pd.read_excel(r"Atividade-29.10.xlsx", 
sheet_name = "Página1",
skiprows=2,
)

df.columns = df.columns.str.strip()

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


df['Valor Total'] = df['Valor Total'].astype(str).str.replace('R$', '', regex=False).str.replace(',', '.', regex=False).str.strip()
df['Valor Total'] = pd.to_numeric(df['Valor Total'], errors='coerce')


print(df)


print("")
print("")

print("\n=======================================================")
print("                 O produto mais vendido                  ")
print("=======================================================")


receita_por_produto = df.groupby('Produto vendido')['Valor Total'].sum().sort_values(ascending=False)

print(receita_por_produto.to_string())


print("")
print("")

print("\n=======================================================")
print("           Maior valor vendido por vendedor"             )
print("=======================================================")

receita_por_vendedor = df.groupby('Vendedor')['Valor Total'].sum().sort_values(ascending=False)

print(receita_por_vendedor.to_string())