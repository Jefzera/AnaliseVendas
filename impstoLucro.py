import pandas as pd

# Carregar os dados
dados = pd.read_excel('base_vendas.xlsx')

# Garantir que a coluna 'data_venda' esteja no formato de data
dados['data_venda'] = pd.to_datetime(dados['data_venda'], errors='coerce')

# Filtrar os dados para o ano de 2023
dados_2023 = dados[dados['data_venda'].dt.year == 2023]

# Calcular o lucro como diferença entre valor_vendido e custo
dados_2023['lucro'] = dados_2023['valor_vendido'] - dados_2023['custo']

# Agrupar os dados por mês e calcular a soma de custos, impostos, lucro e vendas
resumo_mensal = dados_2023.groupby(dados_2023['data_venda'].dt.month).agg(
    total_impostos=('impostos', 'sum'),
    total_custo=('custo', 'sum'),
    total_lucro=('lucro', 'sum'),
    total_vendas=('valor_vendido', 'sum')
).reset_index()

# Renomear as colunas para exibição
resumo_mensal.columns = ['Mês', 'Total de Impostos (R$)', 'Total de Custo (R$)', 'Total de Lucro (R$)', 'Total de Vendas (R$)']

# Exportar os resultados para um arquivo Excel
resumo_mensal.to_excel('analise_impostos_custo_lucro_2023.xlsx', index=False)

# Confirmar que o arquivo foi salvo
print("Arquivo Excel 'analise_impostos_custo_lucro_2023.xlsx' foi salvo com sucesso.")
