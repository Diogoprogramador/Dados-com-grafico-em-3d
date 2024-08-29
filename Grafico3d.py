import openpyxl
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
import pandas as pd
df = pd.read_excel("vendas.xlsx")

print(df)

print(df['Valor Final'].sum())

valor_final_por_produto = df.groupby('Produto')['Valor Final'].sum().reset_index()


print(valor_final_por_produto)

valor_final_por_produto_loja = df.groupby(['Produto', 'ID Loja'])['Valor Final'].sum().reset_index()


print(valor_final_por_produto_loja)


# Função para ler dados da planilha
def ler_dados_excel(caminho_arquivo):
    workbook = openpyxl.load_workbook(caminho_arquivo)
    sheet = workbook.active
    dados = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        dados.append(row)

    return dados


# Ler dados
caminho_arquivo = 'Vendas.xlsx'
dados = ler_dados_excel(caminho_arquivo)

# Preparar os dados para o gráfico
x = [row[0] for row in dados]  # IDs ou categorias
y = [row[1] for row in dados]  # Valores de uma coluna
z = [0] * len(x)  # Z-axial com valor inicial de 0

fig = plt.figure()
ax = fig.add_subplot(111, projection='3d')
ax.bar(x, y, zs=z, zdir='y', alpha=0.8)

ax.set_xlabel('Categorias')
ax.set_ylabel('Valores')
ax.set_zlabel('Z')

plt.show()
