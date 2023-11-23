import pyodbc
import pandas as pd
import ctypes

# Parâmetros de conexão com o banco de dados SQL Server
server = 'SVRERP,1433'
database = 'PROTHEUS12_R27'
username = 'coognicao'
password = '0705@Abc'
driver = '{ODBC Driver 17 for SQL Server}'

# Caminho para o arquivo Excel
excel_file_path = r'\\192.175.175.4\f\INTEGRANTES\ELIEZER\PROJETO SOLIDWORKS TOTVS\M-048-020-284.xlsx'

# Tente estabelecer a conexão com o banco de dados
try:
    conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
    
    # Cria um cursor para executar comandos SQL
    cursor = conn.cursor()

    # Query SELECT
    select_query = "SELECT * FROM PROTHEUS12_R27.dbo.SG1010 WHERE G1_COD = 'M-048-020-284'"

    # Executa a query SELECT e obtém os resultados em um DataFrame
    df_sql = pd.read_sql(select_query, conn)
    
    # Remove espaços em branco da coluna 'G1_COD'
    df_sql['G1_COMP'] = df_sql['G1_COMP'].str.strip()

    # Carrega a planilha do Excel em um DataFrame
    df_excel = pd.read_excel(excel_file_path, sheet_name='Planilha1')

    # Verifica se há alguma correspondência entre as duas colunas
    result = df_excel.iloc[:, 1].isin(df_sql['G1_COMP']).any()

    # Exibe uma caixa de diálogo com base no resultado
    if result:
        ctypes.windll.user32.MessageBoxW(0, "Encontrado um valor correspondente na planilha!", "Sucesso", 1)
    else:
        ctypes.windll.user32.MessageBoxW(0, "Nenhum valor correspondente encontrado na planilha.", "Sucesso", 1)

except pyodbc.Error as ex:
    # Exibe uma caixa de diálogo se a conexão ou a consulta falhar
    ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou consulta. Erro: {str(ex)}", "Erro", 0)

finally:
    # Fecha a conexão com o banco de dados
    conn.close()
