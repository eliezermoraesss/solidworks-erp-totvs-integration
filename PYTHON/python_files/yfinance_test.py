import yfinance as yf

# Nome do ativo (exemplo: "PETR4.SA" para Petrobras ou "HGLG11.SA" para um FII)
ticker = "AAPL"

# Baixar os dados históricos do ativo
ativo = yf.Ticker(ticker)
dados_historicos = ativo.history(period="3mo")  # Pega os dados do último mês

# Imprimir os dados históricos
print(dados_historicos)

# Obter informações gerais do ativo
info_ativo = ativo.info
print(info_ativo)
