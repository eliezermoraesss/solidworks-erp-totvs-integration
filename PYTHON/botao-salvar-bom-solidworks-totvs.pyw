import pyodbc
import pandas as pd
import ctypes
import os
import re

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
        
def validar_formato_codigos(df_excel, posicao_coluna_codigo):
    # Valida os códigos do Excel

    # Formatos de código permitidos
    formatos_codigo = [
        r'^(C|M)\-\d{3}\-\d{3}\-\d{3}$',
        r'^(E\d{4}\-\d{3}\-\d{3})$',
    ]

    # Valida cada código
    validacao_codigos = df_excel.iloc[:, posicao_coluna_codigo].apply(lambda x: any(re.match(formato, str(x)) for formato in formatos_codigo))

    return validacao_codigos

def verificar_codigo_repetido(df_excel):
    codigos_repetidos = df_excel.iloc[:, posicao_coluna_codigo_excel][df_excel.iloc[:, posicao_coluna_codigo_excel].duplicated()]
    
    return codigos_repetidos

def verificar_cadastro_produtos(codigos):
    codigos_sem_cadastro = []
    for codigo_produto in codigos:
        query_consulta_produto = f"""
        SELECT B1_COD FROM PROTHEUS12_R27.dbo.SB1010 WHERE B1_COD = '{codigo_produto}';
        """
        cursor.execute(query_consulta_produto)
        resultado = cursor.fetchone()

        if not resultado:
            codigos_sem_cadastro.append(codigo_produto)

    return codigos_sem_cadastro
    
nome_desenho = ler_variavel_ambiente_codigo_desenho()
print(nome_desenho)

base_path = os.environ.get('TEMP')
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

    posicao_coluna_codigo_excel = 1
    posicao_coluna_descricao_excel = 2
    posicao_coluna_quantidade_excel = 3
    
    # Carrega a planilha do Excel em um DataFrame
    df_excel = pd.read_excel(excel_file_path, sheet_name='Planilha1', header=None)
    
    # Exclui a última linha do DataFrame
    df_excel = df_excel.drop(df_excel.index[-1])

    print(df_excel)

    valid_codigos = validar_formato_codigos(df_excel, posicao_coluna_codigo_excel)

    valid_quantidades = df_excel.iloc[:, posicao_coluna_quantidade_excel].notna() & (df_excel.iloc[:, posicao_coluna_quantidade_excel] != '') & (pd.to_numeric(df_excel.iloc[:, posicao_coluna_quantidade_excel], errors='coerce') > 0)

    # Exibe uma mensagem de erro se os códigos ou quantidades não estiverem no formato esperado
    if not valid_codigos.all():
        ctypes.windll.user32.MessageBoxW(0, "Códigos inválidos encontrados. Corrija os códigos no formato correto.", "Erro", 0)

    if not valid_quantidades.all():
        ctypes.windll.user32.MessageBoxW(0, "Quantidades inválidas encontradas. As quantidades devem ser números, não nulas, sem espaços em branco e maiores que zero.", "Erro", 0)

    if valid_codigos.all() and valid_quantidades.all():

        # Remove espaços em branco da coluna número 2
        df_excel.iloc[:, posicao_coluna_codigo_excel] = df_excel.iloc[:, posicao_coluna_codigo_excel].str.strip()

        # Encontra códigos que são iguais entre SQL e Excel
        codigos_em_comum = df_sql['G1_COMP'].loc[df_sql['G1_COMP'].isin(df_excel.iloc[:, posicao_coluna_codigo_excel])].tolist()
        codigos_adicionados_bom = df_excel.iloc[:, posicao_coluna_codigo_excel].loc[~df_excel.iloc[:, posicao_coluna_codigo_excel].isin(df_sql['G1_COMP'])].tolist()
        codigos_removidos_bom = df_sql['G1_COMP'].loc[~df_sql['G1_COMP'].isin(df_excel.iloc[:, posicao_coluna_codigo_excel])].tolist()

        # Exibe uma caixa de diálogo com base nos resultados
        if codigos_em_comum:
            ctypes.windll.user32.MessageBoxW(0, f"Códigos em comuns: {codigos_em_comum}", "ITENS EM COMUM", 1)

        if codigos_adicionados_bom:
            ctypes.windll.user32.MessageBoxW(0, f"Itens adicionados: {codigos_adicionados_bom}", "ITENS ADICIONADOS", 1)

        if codigos_removidos_bom:
            ctypes.windll.user32.MessageBoxW(0, f"Itens removidos: {codigos_removidos_bom}", "ITENS REMOVIDOS", 1)
            
        # Verificar cadastro dos códigos em comum e adicionados
        codigos_sem_cadastro = verificar_cadastro_produtos(codigos_em_comum + codigos_adicionados_bom)
        
        if codigos_sem_cadastro:
            mensagem = f"Os seguintes itens não possuem cadastro:\n\n{', '.join(codigos_sem_cadastro)}\n\nEfetue o cadastro e tente novamente."
            ctypes.windll.user32.MessageBoxW(0, mensagem, "Códigos sem Cadastro", 1)
            
        codigos_repetidos = verificar_codigo_repetido(df_excel)
    
        # Exibe uma mensagem se houver códigos repetidos
        if not codigos_repetidos.empty:
            raise ValueError("Códigos repetidos encontrados.")  
            
except pyodbc.Error as ex:
    # Exibe uma caixa de diálogo se a conexão ou a consulta falhar
    ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou consulta. Erro: {str(ex)}", "Erro", 0)

except ValueError as ve:
    ctypes.windll.user32.MessageBoxW(0, f"Códigos repetidos encontrados: {codigos_repetidos.tolist()}", "Aviso", 0)

finally:
    # Fecha a conexão com o banco de dados
    conn.close()
    delete_file_if_exists(excel_file_path)
