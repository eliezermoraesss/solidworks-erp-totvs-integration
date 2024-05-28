import pandas as pd
import pyodbc  # Ou outra biblioteca de conexão com o banco de dados que você esteja usando

server = 'SVRERP,1433'
database = 'PROTHEUS12_R27' # PROTHEUS12_R27 (base de produção) PROTHEUS1233_HML (base de desenvolvimento/teste)
username = 'coognicao'
password = '0705@Abc'
driver = '{ODBC Driver 17 for SQL Server}'
conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')

def consultar_filhos_iterativo(df, codigo, df_mp):
    stack = [codigo]  # Inicializa a pilha com o código inicial
    
    while stack:
        codigo_atual = stack.pop()  # Pega o código atual da pilha
        
        # Consulta os filhos do código atual
        filhos = df[df['CÓDIGO'].str.startswith(codigo_atual)]
        
        # Filtra os filhos do tipo MP e adiciona ao DataFrame df_mp
        filhos_mp = filhos[filhos['TIPO'] == 'MP']
        df_mp = pd.concat([df_mp, filhos_mp], ignore_index=True)
        
        # Adiciona os códigos dos filhos à pilha para continuar a busca
        for filho in filhos['CÓDIGO']:
            stack.append(filho)
    
    return df_mp

# Definir a query SQL
query = '''
SELECT struct.G1_COMP AS "CÓDIGO", 
       prod.B1_DESC AS "DESCRIÇÃO", 
       struct.G1_QUANT AS "QUANT.", 
       struct.G1_XUM AS "UNID. MED.", 
       prod.B1_UCOM AS "ULT. ATUALIZ.", 
       prod.B1_TIPO AS "TIPO", 
       prod.B1_LOCPAD AS "ARMAZÉM", 
       FORMAT(prod.B1_UPRC, 'N2') AS "VALOR UNIT. (R$)", 
       FORMAT(G1_QUANT * B1_UPRC, 'N2') AS "SUB-TOTAL (R$)" 
FROM SG1010 struct 
INNER JOIN SB1010 prod ON struct.G1_COMP = prod.B1_COD 
WHERE struct.G1_COD = 'M-042-007-900' 
      AND G1_REVFIM = (SELECT MAX(G1_REVFIM) 
                       FROM SG1010 
                       WHERE G1_COD = 'M-042-007-900' 
                             AND G1_REVFIM <> 'ZZZ' 
                             AND D_E_L_E_T_ <> '*') 
      AND struct.G1_REVFIM <> 'ZZZ' 
      AND struct.D_E_L_E_T_ <> '*' 
ORDER BY prod.B1_COD ASC;
'''

# Executar a query e salvar o resultado em um DataFrame
df = pd.read_sql(query, conn)

# Fechar a conexão com o banco de dados
conn.close()

# Exibir o DataFrame
print(df)

# Filtrar os itens do tipo MP
df_mp = df[df['TIPO'] == 'MP']

# Filtrar os itens do tipo PI ou PA
df_pi_pa = df[(df['TIPO'] == 'PI') | (df['TIPO'] == 'PA')]

# Exibir os DataFrames separados
print("Itens do tipo MP:")
print(df_mp)

print("\nItens do tipo PI ou PA:")
print(df_pi_pa)



# Inicializa o DataFrame para os itens do tipo MP
df_mp_encontrado = pd.DataFrame(columns=df.columns)

# Consulta os itens do tipo MP a partir dos itens do tipo PI ou PA
for index, row in df_pi_pa.iterrows():
    df_mp_encontrado = consultar_filhos_iterativo(df, row['CÓDIGO'], df_mp_encontrado)

# Exibe os MP encontrados
print("Itens do tipo MP encontrados:")
print(df_mp_encontrado)
