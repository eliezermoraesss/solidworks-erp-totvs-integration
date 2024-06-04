import pandas as pd
import pyodbc

server = 'SVRERP,1433'
database = 'PROTHEUS12_R27' # PROTHEUS12_R27 (base de produção) PROTHEUS1233_HML (base de desenvolvimento/teste)
username = 'coognicao'
password = '0705@Abc'
driver = '{ODBC Driver 17 for SQL Server}'

codigo = 'M-059-005-610'

conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')

# Definir a query
query = f"""
DECLARE @CodigoPai VARCHAR(50) = '{codigo}'; -- Substitua pelo código pai que deseja consultar

-- CTE para selecionar os itens pai e seus subitens recursivamente
WITH ListMP AS (
    -- Selecionar o item pai inicialmente
    SELECT G1_COD AS "CÓDIGO", G1_COMP AS "COMPONENTE", 0 AS Nivel
    FROM SG1010
    WHERE G1_COD = @CodigoPai AND G1_REVFIM = (
        SELECT MAX(G1_REVFIM) 
        FROM SG1010 
        WHERE G1_COD = @CodigoPai AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*'
    ) AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*'

    UNION ALL

    -- Selecione os subitens de cada item pai
    SELECT sub.G1_COD, sub.G1_COMP, pai.Nivel + 1
    FROM SG1010 AS sub
    INNER JOIN ListMP AS pai ON sub.G1_COD = pai."COMPONENTE"
    WHERE pai.Nivel < 100 -- Defina o limite máximo de recursão aqui
    AND sub.G1_REVFIM <> 'ZZZ' AND sub.D_E_L_E_T_ <> '*'
)

-- Selecione todas as matérias-primas (tipo = 'MP') que correspondem aos itens encontrados
SELECT DISTINCT
    mat.G1_COD AS "CODIGO PAI",
    mat.G1_COMP AS "CÓDIGO", 
    prod.B1_DESC AS "DESCRIÇÃO", 
    mat.G1_QUANT AS "QUANT.", 
    mat.G1_XUM AS "UNID. MED.", 
    prod.B1_UCOM AS "ULT. ATUALIZ.",
    prod.B1_TIPO AS "TIPO", 
    prod.B1_LOCPAD AS "ARMAZÉM", 
    prod.B1_UPRC AS "VALOR UNIT. (R$)",
    (G1_QUANT * B1_UPRC) AS "VALOR TOTAL (R$)"
FROM SG1010 AS mat
INNER JOIN ListMP AS pai ON mat.G1_COD = pai."CÓDIGO"
INNER JOIN SB1010 AS prod ON mat.G1_COMP = prod.B1_COD
WHERE prod.B1_TIPO = 'MP'
AND prod.B1_LOCPAD IN ('01','03')
AND mat.G1_REVFIM <> 'ZZZ' 
AND mat.D_E_L_E_T_ <> '*'
ORDER BY mat.G1_COMP ASC;
"""

# Executar a query e salvar os resultados em um DataFrame
df = pd.read_sql(query, conn)

# Consolidar a quantidade conforme os códigos duplicados
consolidated_df = df.groupby('CÓDIGO').agg({
    'DESCRIÇÃO': 'first',
    'QUANT.': 'sum',
    'UNID. MED.': 'first',
    'ULT. ATUALIZ.': 'first',
    'TIPO': 'first',
    'ARMAZÉM': 'first',
    'VALOR UNIT. (R$)': 'first',
    'VALOR TOTAL (R$)': 'sum'
}).reset_index()

# Fechar a conexão com o banco de dados
conn.close()

# Salvar o DataFrame consolidado em um arquivo Excel
output_file = f'{codigo}_MP_POR_ESTRUTURA.xlsx'
consolidated_df.to_excel(output_file, index=False, engine='openpyxl')

# Mostrar o DataFrame consolidado
print(consolidated_df)
''