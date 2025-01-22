import pandas as pd

# Carregar os dados
dados = pd.read_excel('base_vendas.xlsx')

# Garantir que a coluna 'data_venda' esteja no formato de data
dados['data_venda'] = pd.to_datetime(dados['data_venda'], errors='coerce')

# Filtrar os dados para o ano de 2023
dados_2023 = dados[dados['data_venda'].dt.year == 2023]

# Agrupar os dados por mês, id_produto e somar o valor vendido
vendas_por_mes_produto = dados_2023.groupby([dados_2023['data_venda'].dt.month, 'id_produto'])['valor_vendido'].sum()

# Resetando o índice para manipulação
vendas_por_mes_produto = vendas_por_mes_produto.reset_index()

# Preparar um DataFrame para armazenar os 5 produtos mais e menos vendidos por mês
resultados = []

# Para cada mês (1 a 12)
for mes in range(1, 13):
    vendas_mes = vendas_por_mes_produto[vendas_por_mes_produto['data_venda'] == mes]
    
    # Selecionar os 5 produtos mais vendidos (do maior para o menor)
    mais_vendidos = vendas_mes.nlargest(5, 'valor_vendido')
    
    # Selecionar os 5 produtos menos vendidos (do menor para o maior)
    menos_vendidos = vendas_mes.nsmallest(5, 'valor_vendido')
    
    # Adicionar os resultados
    mais_vendidos['tipo'] = 'Mais Vendidos'
    menos_vendidos['tipo'] = 'Menos Vendidos'
    
    resultados.append(mais_vendidos)
    resultados.append(menos_vendidos)

# Concatenar todos os resultados
resultados_df = pd.concat(resultados)

# Renomear as colunas para exibição
resultados_df = resultados_df[['data_venda', 'id_produto', 'valor_vendido', 'tipo']]
resultados_df.columns = ['Mês', 'ID Produto', 'Valor Vendido (R$)', 'Tipo']

# Exportar os resultados para um arquivo Excel
resultados_df.to_excel('produtos_mais_menos_vendidos_2023.xlsx', index=False)

# Confirmar que o arquivo foi salvo
print("Arquivo Excel 'produtos_mais_menos_vendidos_2023.xlsx' foi salvo com sucesso.")
