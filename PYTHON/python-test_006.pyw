import pandas as pd

# Especifique o caminho do arquivo Excel
caminho_do_arquivo = r'\\192.175.175.4\f\INTEGRANTES\ELIEZER\PROJETO SOLIDWORKS TOTVS\M-048-020-284.xlsx'

# Carregue a planilha Excel em um DataFrame do pandas
df = pd.read_excel(caminho_do_arquivo, header=None)

# Inverta as linhas do DataFrame
df_invertido = df[::-1]

# Especifique o caminho para salvar a nova planilha invertida
caminho_saida = r'\\192.175.175.4\f\INTEGRANTES\ELIEZER\PROJETO SOLIDWORKS TOTVS\M-048-020-284_invertida.xlsx'

# Salve a planilha invertida em um novo arquivo Excel
df_invertido.to_excel(caminho_saida, header=None, index=False)

print(f'A planilha invertida foi salva em {caminho_saida}')

