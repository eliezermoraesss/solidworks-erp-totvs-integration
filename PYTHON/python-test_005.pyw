import pyodbc
import pandas as pd
import ctypes
import os

# Parâmetros de conexão com o banco de dados SQL Server
server = 'SVRERP,1433'
database = 'PROTHEUS12_R27'
username = 'coognicao'
password = '0705@Abc'
driver = '{ODBC Driver 17 for SQL Server}'

def ler_variavel_ambiente_codigo_desenho():
    # Recupera o valor da variável de ambiente
    return os.getenv('CODIGO_DESENHO')

def delete_file_if_exists(excel_file_path):
    if os.path.exists(excel_file_path):
        os.remove(excel_file_path)       
    
nome_desenho = ler_variavel_ambiente_codigo_desenho()
print(nome_desenho)

base_path = r'\\192.175.175.4\f\INTEGRANTES\ELIEZER\PROJETO SOLIDWORKS TOTVS'
excel_file_path = os.path.join(base_path, nome_desenho + '.xlsx')

# Arrays para armazenar os códigos
codigos_adicionados_bom = [] # ITENS ADICIONADOS
codigos_removidos_bom = [] # ITENS REMOVIDOS
codigos_em_comum = [] # ITENS EM COMUM

# Tente estabelecer a conexão com o banco de dados
try:
    conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
    
    # Cria um cursor para executar comandos SQL
    cursor = conn.cursor()

    # Query SELECT
    select_query = f"""SELECT * FROM PROTHEUS12_R27.dbo.SG1010 WHERE G1_COD = '{nome_desenho}'
        AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*'
        AND G1_REVFIM = (SELECT MAX(G1_REVFIM) FROM PROTHEUS12_R27.dbo.SG1010 WHERE G1_COD = '{nome_desenho}' AND G1_REVFIM <> 'ZZZ');
    """

    # Executa a query SELECT e obtém os resultados em um DataFrame
    df_sql = pd.read_sql(select_query, conn)

    # Remove espaços em branco da coluna 'G1_COD'
    df_sql['G1_COMP'] = df_sql['G1_COMP'].str.strip()

    # Carrega a planilha do Excel em um DataFrame e inverte as linhas
    df_excel = pd.read_excel(excel_file_path, sheet_name='Planilha1', header=None)

    print(df_excel)

    # Obtém a posição da coluna número 2 no DataFrame do Excel
    posicao_coluna_excel = 1  # A coluna número 2 tem índice 1 (índices começam do 0)

    # Remove espaços em branco da coluna número 2
    df_excel.iloc[:, posicao_coluna_excel] = df_excel.iloc[:, posicao_coluna_excel].str.strip()

    # Encontra códigos que são iguais entre SQL e Excel
    codigos_em_comum = df_sql['G1_COMP'].loc[df_sql['G1_COMP'].isin(df_excel.iloc[:, posicao_coluna_excel])].tolist()
    codigos_adicionados_bom = df_excel.iloc[:, posicao_coluna_excel].loc[~df_excel.iloc[:, posicao_coluna_excel].isin(df_sql['G1_COMP'])].tolist()
    codigos_removidos_bom = df_sql['G1_COMP'].loc[~df_sql['G1_COMP'].isin(df_excel.iloc[:, posicao_coluna_excel])].tolist()

    # Exibe uma caixa de diálogo com base nos resultados
    if codigos_em_comum:
        ctypes.windll.user32.MessageBoxW(0, f"Códigos em comuns: {codigos_em_comum}", "ITENS EM COMUM", 1)

    if codigos_adicionados_bom:
        ctypes.windll.user32.MessageBoxW(0, f"Itens adicionados: {codigos_adicionados_bom}", "ITENS ADICIONADOS", 1)

    if codigos_removidos_bom:
        ctypes.windll.user32.MessageBoxW(0, f"Itens removidos: {codigos_removidos_bom}", "ITENS REMOVIDOS", 1)

except pyodbc.Error as ex:
    # Exibe uma caixa de diálogo se a conexão ou a consulta falhar
    ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou consulta. Erro: {str(ex)}", "Erro", 0)

finally:
    # Fecha a conexão com o banco de dados
    conn.close()
    delete_file_if_exists(excel_file_path)
