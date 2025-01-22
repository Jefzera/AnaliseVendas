import pandas as pd

# Carregar os dados
dados = pd.read_excel('base_vendas.xlsx')

# Garantir que a coluna 'data_venda' esteja no formato de data
dados['data_venda'] = pd.to_datetime(dados['data_venda'], errors='coerce')

# Filtrar os dados para o ano de 2023
dados_2023 = dados[dados['data_venda'].dt.year == 2023]

# Agrupar os dados por mês e somar o valor das vendas
vendas_por_mes = dados_2023.groupby(dados_2023['data_venda'].dt.month)['valor_vendido'].sum()

# Converter o resultado para um DataFrame para exportar facilmente
vendas_por_mes_df = vendas_por_mes.reset_index()
vendas_por_mes_df.columns = ['Mês', 'Total Vendas (R$)']  # Renomear as colunas

# Exportar o resultado para um arquivo Excel
vendas_por_mes_df.to_excel('vendas_por_mes_2023.xlsx', index=False)

# Confirmar que o arquivo foi salvo
print("Arquivo Excel 'vendas_por_mes_2023.xlsx' foi salvo com sucesso.")