import pandas as pd


# Carregar os dados
dados = pd.read_excel('base_vendas.xlsx')

# Garantir que a coluna 'data_venda' esteja no formato de data
dados['data_venda'] = pd.to_datetime(dados['data_venda'], errors='coerce')

# Filtrar os dados para o ano de 2023
dados_2023 = dados[dados['data_venda'].dt.year == 2023]

# Criar uma coluna com o nome do vendedor (pode ser id_cliente ou nome completo)
dados_2023['vendedor'] = dados_2023['primeiro_nome'] + ' ' + dados_2023['sobrenome']

# Agrupar os dados por mês e por vendedor, somando o valor das vendas
vendas_por_mes_vendedor = dados_2023.groupby([dados_2023['data_venda'].dt.month, 'vendedor'])['valor_vendido'].sum()

# Resetando o índice para manipulação
vendas_por_mes_vendedor = vendas_por_mes_vendedor.reset_index()

# Preparar um DataFrame para armazenar os 5 melhores e 5 piores vendedores por mês
resultados = []

# Para cada mês (1 a 12)
for mes in range(1, 13):
    vendas_mes = vendas_por_mes_vendedor[vendas_por_mes_vendedor['data_venda'] == mes]
    
    # Ordenar os vendedores por vendas (do maior para o menor)
    melhores_vendedores = vendas_mes.nlargest(5, 'valor_vendido')
    piores_vendedores = vendas_mes.nsmallest(5, 'valor_vendido')
    
    # Adicionar ao resultado
    melhores_vendedores['tipo'] = 'Melhores'
    piores_vendedores['tipo'] = 'Piores'
    
    resultados.append(melhores_vendedores)
    resultados.append(piores_vendedores)

# Concatenar todos os resultados
resultados_df = pd.concat(resultados)

# Renomear as colunas para exibição
resultados_df = resultados_df[['data_venda', 'vendedor', 'valor_vendido', 'tipo']]
resultados_df.columns = ['Mês', 'Vendedor', 'Total Vendas (R$)', 'Tipo']

# Exportar os resultados para um arquivo Excel
resultados_df.to_excel('melhores_piores_vendedores_2023.xlsx', index=False)

# Confirmar que o arquivo foi salvo
print("Arquivo Excel 'melhores_piores_vendedores_2023.xlsx' foi salvo com sucesso.")