import pyodbc
import pandas as pd
import ctypes

# Parâmetros de conexão com o banco de dados SQL Server
server = 'SVRERP,1433'
database = 'PROTHEUS12_R27'
username = 'coognicao'
password = '0705@Abc'
driver = '{ODBC Driver 17 for SQL Server}'

# Caminho para o arquivo Excel (caminho bruto)
excel_file_path = r'\\192.175.175.4\f\INTEGRANTES\ELIEZER\PROJETO SOLIDWORKS TOTVS\M-048-020-284.xlsx'

# Arrays para armazenar os códigos
codigos_excel = []
codigos_sql = []

# Tente estabelecer a conexão com o banco de dados
try:
    conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
    
    # Cria um cursor para executar comandos SQL
    cursor = conn.cursor()

    # Query SELECT
    select_query = "SELECT * FROM PROTHEUS12_R27.dbo.SG1010 WHERE G1_COD = 'M-048-020-284' AND D_E_L_E_T_ <> '*'"

    # Executa a query SELECT e obtém os resultados em um DataFrame
    df_sql = pd.read_sql(select_query, conn)

    # Remove espaços em branco da coluna 'G1_COD'
    df_sql['G1_COMP'] = df_sql['G1_COMP'].str.strip()

    # Carrega a planilha do Excel em um DataFrame
    df_excel = pd.read_excel(excel_file_path, sheet_name='Planilha1')

    # Obtém a posição da coluna número 2 no DataFrame do Excel
    posicao_coluna_excel = 1  # A coluna número 2 tem índice 1 (índices começam do 0)

    # Remove espaços em branco da coluna número 2
    df_excel.iloc[:, posicao_coluna_excel] = df_excel.iloc[:, posicao_coluna_excel].str.strip()

    # Encontra códigos que são iguais entre SQL e Excel
    codigos_comuns = df_sql['G1_COMP'].loc[df_sql['G1_COMP'].isin(df_excel.iloc[:, posicao_coluna_excel])].tolist()
    codigos_excel = df_excel.iloc[:, posicao_coluna_excel].loc[~df_excel.iloc[:, posicao_coluna_excel].isin(df_sql['G1_COMP'])].tolist()
    codigos_sql = df_sql['G1_COMP'].loc[~df_sql['G1_COMP'].isin(df_excel.iloc[:, posicao_coluna_excel])].tolist()

    # Exibe uma caixa de diálogo com base nos resultados
    if codigos_comuns:
        ctypes.windll.user32.MessageBoxW(0, f"Códigos comuns: {codigos_comuns}", "Códigos comuns", 1)

    if codigos_excel:
        ctypes.windll.user32.MessageBoxW(0, f"Códigos no Excel, mas não no SQL: {codigos_excel}", "Excel > SQL", 1)

    if codigos_sql:
        ctypes.windll.user32.MessageBoxW(0, f"Códigos no SQL, mas não no Excel: {codigos_sql}", "SQL > Excel", 1)

except pyodbc.Error as ex:
    # Exibe uma caixa de diálogo se a conexão ou a consulta falhar
    ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou consulta. Erro: {str(ex)}", "Erro", 0)

finally:
    # Fecha a conexão com o banco de dados
    conn.close()
