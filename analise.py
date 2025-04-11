import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Leitura das planilhas
vendas = pd.read_csv("vendas.xlsx")
estoque = pd.read_csv("estoque.xlsx")
funcionarios = pd.read_csv("funcionarios.xlsx")

# Cálculo do estoque atual
estoque["Estoque Atual"] = estoque["Estoque Inicial"] + estoque["Entradas"] - estoque["Saídas"]

# Total de vendas por produto
vendas_produto = vendas.groupby("Produto")["Quantidade"].sum().sort_values(ascending=False)
print("Total vendido por produto:\n", vendas_produto)

# Faturamento por região
vendas["Faturamento"] = vendas["Quantidade"] * vendas["Preço"]
faturamento_regiao = vendas.groupby("Região")["Faturamento"].sum()
print("\nFaturamento por região:\n", faturamento_regiao)

# Produto mais lucrativo
lucro_produto = vendas.groupby("Produto")["Lucro"].sum().sort_values(ascending=False)
print("\nProduto com maior lucro:\n", lucro_produto)

# Produtos com estoque baixo
estoque_baixo = estoque[estoque["Estoque Atual"] < 10]
print("\nProdutos com estoque baixo:\n", estoque_baixo)

# Funcionários mais antigos (top 3)
experientes = funcionarios.sort_values(by="Tempo de Empresa", ascending=False).head(3)
print("\nFuncionários mais antigos:\n", experientes)

# Média salarial por setor
salario_por_setor = funcionarios.groupby("Setor")["Salário"].mean()
print("\nMédia salarial por setor:\n", salario_por_setor)

# Gráfico de barras dos produtos mais vendidos
plt.figure(figsize=(8, 4))
sns.barplot(x=vendas_produto.values, y=vendas_produto.index, palette="viridis")
plt.title("Produtos Mais Vendidos")
plt.xlabel("Quantidade Vendida")
plt.tight_layout()
plt.show()

# Faturamento por região (gráfico de pizza)
faturamento_regiao.plot(kind="pie", autopct="%1.1f%%", figsize=(6, 6), title="Faturamento por Região")
plt.ylabel("")  # remove o rótulo
plt.show()

# Exportar relatório consolidado
with pd.ExcelWriter("relatorio_final.xlsx", engine="openpyxl") as writer:
    vendas_produto.to_frame(name="Qtd Vendida").to_excel(writer, sheet_name="Vendas")
    faturamento_regiao.to_frame(name="Faturamento").to_excel(writer, sheet_name="Faturamento")
    lucro_produto.to_frame(name="Lucro Total").to_excel(writer, sheet_name="Lucro")
    estoque.to_excel(writer, sheet_name="Estoque")
    salario_por_setor.to_frame(name="Média Salarial").to_excel(writer, sheet_name="RH")
    experientes.to_excel(writer, sheet_name="Funcionários Experientes")
