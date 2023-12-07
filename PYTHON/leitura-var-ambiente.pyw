import os
import ctypes

# Recupera o valor da variável de ambiente
nome_arquivo = os.getenv('CODIGO_DESENHO')

# Se a variável de ambiente não existir, forneça um valor padrão ou trate a situação conforme necessário
if nome_arquivo is None:
    nome_arquivo = "valor_padrao"

# Agora você pode usar a variável nome_arquivo no seu script
print(f"Código do desenho: {nome_arquivo}")

ctypes.windll.user32.MessageBoxW(0, f"env: {nome_arquivo}", "Var. Ambiente", 1)
