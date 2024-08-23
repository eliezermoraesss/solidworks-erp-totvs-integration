import investpy

# Nome do ativo (exemplo: "PETR4" para Petrobras ou "HGLG11" para um FII)
ticker = "HGLG11"

# Pegar dados históricos de um FII ou ação
dados_historicos = investpy.get_stock_historical_data(stock=ticker, country='brazil', from_date='01/01/2024', to_date='22/08/2024')

# Imprimir os dados históricos
print(dados_historicos)

# Pegar informações detalhadas do ativo
info_ativo = investpy.get_stock_information(stock=ticker, country='brazil')
print(info_ativo)
