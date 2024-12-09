import pandas as pd

# Simulando o dataframe com base na descrição
data = {
    "CÓDIGO": [
        (2, "COMERCIAL"), (3, "MP"), (4, "PROD. COMER. IMPORT. DIR."),
        (5, "MAT. PRIMA IMPORTADA"), (6, "TRAT. SUPERF."), (8, "TOTAL GERAL"),
        (10, "FATOR ENAPLIC"), (11, "SUGESTÃO DE VENDA"), (12, "PREÇO DO DÓLAR"), (13, "VENDA EM DÓLAR")
    ],
    "DESCRIÇÃO": [0, 81.6, 0, 0, 0, 81.6, 3.75, 306, 5.2, 58.84615],
    "QUANT.": [0, 9.6, 0, 0, 0, 0, None, None, None, None],
}

df = pd.DataFrame(data)

# Separar o primeiro dataframe (até "TOTAL GERAL")
idx_total_geral = df[df["CÓDIGO"].apply(lambda x: x[1] == "TOTAL GERAL")].index[0]
df1 = df.iloc[: idx_total_geral + 1]  # Inclui até "TOTAL GERAL"

# Separar o segundo dataframe (de "FATOR ENAPLIC" até "VENDA EM DÓLAR")
idx_fator_enaplic = df[df["CÓDIGO"].apply(lambda x: x[1] == "FATOR ENAPLIC")].index[0]
df2 = df.iloc[idx_fator_enaplic:]  # A partir de "FATOR ENAPLIC"

# Remover a coluna "QUANT." do segundo dataframe
df2 = df2.drop(columns=["QUANT."])

# Exibir os dois dataframes
print("Primeiro DataFrame:")
print(df1)

print("\nSegundo DataFrame:")
print(df2)
