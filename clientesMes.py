import pandas as pd

# Carregar os dados
dados = pd.read_excel('base_vendas.xlsx')

# Garantir que a coluna 'data_venda' esteja no formato de data
dados['data_venda'] = pd.to_datetime(dados['data_venda'], errors='coerce')

# Filtrar os dados para o ano de 2023
dados_2023 = dados[dados['data_venda'].dt.year == 2023]

# Agrupar os dados por mês, id_cliente e somar o valor vendido
compras_por_mes_comprador = dados_2023.groupby([dados_2023['data_venda'].dt.month, 'id_cliente'])['valor_vendido'].sum()

# Resetando o índice para manipulação
compras_por_mes_comprador = compras_por_mes_comprador.reset_index()

# Preparar um DataFrame para armazenar os 5 maiores compradores por mês
resultados = []

# Para cada mês (1 a 12)
for mes in range(1, 13):
    compras_mes = compras_por_mes_comprador[compras_por_mes_comprador['data_venda'] == mes]
    
    # Ordenar os compradores por valor comprado (do maior para o menor)
    maiores_compradores = compras_mes.nlargest(5, 'valor_vendido')
    
    # Adicionar ao resultado
    maiores_compradores['tipo'] = 'Maiores Compradores'
    
    resultados.append(maiores_compradores)

# Concatenar todos os resultados
resultados_df = pd.concat(resultados)

# Renomear as colunas para exibição
resultados_df = resultados_df[['data_venda', 'id_cliente', 'valor_vendido', 'tipo']]
resultados_df.columns = ['Mês', 'ID Cliente', 'Valor Total Comprado (R$)', 'Tipo']

# Exportar os resultados para um arquivo Excel
resultados_df.to_excel('maiores_compradores_2023.xlsx', index=False)

# Confirmar que o arquivo foi salvo
print("Arquivo Excel 'maiores_compradores_2023.xlsx' foi salvo com sucesso.")