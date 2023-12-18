import pyodbc
import pandas as pd
import ctypes
import os
import re
from datetime import date

# Parâmetros de conexão com o banco de dados SQL Server
server = 'SVRERP,1433'
database = 'PROTHEUS1233_HML'
username = 'coognicao'
password = '0705@Abc'
driver = '{ODBC Driver 17 for SQL Server}'

# Arrays para armazenar os códigos
codigos_adicionados_bom = []  # ITENS ADICIONADOS
codigos_removidos_bom = []  # ITENS REMOVIDOS
codigos_em_comum = []  # ITENS EM COMUM

posicao_coluna_codigo_excel = 1
posicao_coluna_descricao_excel = 2
posicao_coluna_quantidade_excel = 3

def validar_formato_codigos(df_excel, posicao_coluna_codigo):
    formatos_codigo = [
        r'^(C|M)\-\d{3}\-\d{3}\-\d{3}$',
        r'^(E\d{4}\-\d{3}\-\d{3})$',
    ]

    validacao_codigos = df_excel.iloc[:, posicao_coluna_codigo].apply(
        lambda x: any(re.match(formato, str(x)) for formato in formatos_codigo))

    return validacao_codigos


def ler_variavel_ambiente_codigo_desenho():
    return os.getenv('CODIGO_DESENHO')


def obter_caminho_arquivo_excel(codigo_desenho):
    base_path = os.environ.get('TEMP')
    return os.path.join(base_path, codigo_desenho + '.xlsx')


def delete_file_if_exists(excel_file_path):
    if os.path.exists(excel_file_path):
        os.remove(excel_file_path)


def verificar_codigo_repetido(df_excel):
    
    codigos_repetidos = df_excel.iloc[:, posicao_coluna_codigo_excel][df_excel.iloc[:, posicao_coluna_codigo_excel].duplicated()]
    
    # Exibe uma mensagem se houver códigos repetidos
    if not codigos_repetidos.empty:
        ctypes.windll.user32.MessageBoxW(
            0, f"Códigos repetidos encontrados: {codigos_repetidos.tolist()}", "Aviso", 0)


def verificar_cadastro_produtos(codigos):            
    try:
        conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
        cursor = conn.cursor()

        codigos_sem_cadastro = []
        
        for codigo_produto in codigos:
            query_consulta_produto = f"""
            SELECT B1_COD FROM {database}.dbo.SB1010 WHERE B1_COD = '{codigo_produto}';
            """
            
            cursor.execute(query_consulta_produto)
            resultado = cursor.fetchone()

            if not resultado:
                codigos_sem_cadastro.append(codigo_produto)

        if codigos_sem_cadastro:
                mensagem = f"Os seguintes itens não possuem cadastro:\n\n{', '.join(codigos_sem_cadastro)}\n\nEfetue o cadastro e tente novamente."
                ctypes.windll.user32.MessageBoxW(
                    0, mensagem, "Códigos sem Cadastro", 1)

    except Exception as ex:
        # Exibe uma caixa de diálogo se a conexão ou a consulta falhar
        ctypes.windll.user32.MessageBoxW(
            0, f"Falha na conexão ou consulta. Erro: {str(ex)}", "Erro Verificar Cadastro", 0)

    finally:
        # Fecha a conexão com o banco de dados se estiver aberta
        if 'conn' in locals():
            conn.close()


def remover_linhas_duplicadas_e_consolidar_quantidade(df_excel):
    # Agrupa o DataFrame pela combinação única de código e descrição
    grouped = df_excel.groupby([posicao_coluna_codigo_excel, posicao_coluna_descricao_excel])

    # Inicializa um novo DataFrame para armazenar o resultado
    df_sem_duplicatas = pd.DataFrame(columns=df_excel.columns)

    # Itera sobre os grupos consolidando as quantidades
    for _, group in grouped:
        codigo = group[posicao_coluna_codigo_excel].iloc[0]
        descricao = group[posicao_coluna_descricao_excel].iloc[0]
        quantidade_consolidada = group[posicao_coluna_quantidade_excel].sum()

        # Adiciona uma linha ao DataFrame sem duplicatas
        df_sem_duplicatas = pd.concat([df_sem_duplicatas, group.head(1)])
        df_sem_duplicatas.loc[df_sem_duplicatas.index[-1], posicao_coluna_quantidade_excel] = quantidade_consolidada

    return df_sem_duplicatas


def validacao_de_dados_bom(excel_file_path):

    # Carrega a planilha do Excel em um DataFrame
    df_excel = pd.read_excel(excel_file_path, sheet_name='Planilha1', header=None)

    # Exclui a última linha do DataFrame
    df_excel = df_excel.drop(df_excel.index[-1])

    valid_codigos = validar_formato_codigos(df_excel, posicao_coluna_codigo_excel)

    valid_quantidades = df_excel.iloc[:, posicao_coluna_quantidade_excel].notna(
    ) & (df_excel.iloc[:, posicao_coluna_quantidade_excel] != '') & (pd.to_numeric(df_excel.iloc[:, posicao_coluna_quantidade_excel], errors='coerce') > 0)

    # Exibe uma mensagem de erro se os códigos ou quantidades não estiverem no formato esperado
    if not valid_codigos.all():
        ctypes.windll.user32.MessageBoxW(
            0, "Códigos inválidos encontrados. Corrija os códigos no formato correto.", "Erro", 0)

    if not valid_quantidades.all():
        ctypes.windll.user32.MessageBoxW(
            0, "Quantidades inválidas encontradas. As quantidades devem ser números, não nulas, sem espaços em branco e maiores que zero.", "Erro", 0)

    if valid_codigos.all() and valid_quantidades.all():

        bom_excel_sem_duplicatas = remover_linhas_duplicadas_e_consolidar_quantidade(df_excel)
        bom_excel_sem_duplicatas.iloc[:, posicao_coluna_codigo_excel] = bom_excel_sem_duplicatas.iloc[:, posicao_coluna_codigo_excel].str.strip()
        verificar_cadastro_produtos(bom_excel_sem_duplicatas.iloc[:, posicao_coluna_codigo_excel].tolist())
        verificar_codigo_repetido(bom_excel_sem_duplicatas)
            
    return bom_excel_sem_duplicatas


def verificar_se_existe_estrutura_totvs(nome_desenho):

    query_consulta_estrutura_totvs = f"""SELECT * FROM {database}.dbo.SG1010 WHERE G1_COD = '{nome_desenho}'
        AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*'
        AND G1_REVFIM = (SELECT MAX(G1_REVFIM) FROM {database}.dbo.SG1010 WHERE G1_COD = '{nome_desenho}' AND G1_REVFIM <> 'ZZZ');
    """
    # Tente estabelecer a conexão com o banco de dados
    try:
        conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
        cursor = conn.cursor()

        # Executa a query SELECT e obtém os resultados em um DataFrame
        resultado_query_consulta_estrutura_totvs = pd.read_sql(query_consulta_estrutura_totvs, conn)

        return resultado_query_consulta_estrutura_totvs

    except Exception as ex:
        # Exibe uma caixa de diálogo se a conexão ou a consulta falhar
        ctypes.windll.user32.MessageBoxW(
            0, f"Falha na conexão ou consulta. Erro: {str(ex)}", "Erro Verificar Existe Estrutura", 0)

    finally:
        # Fecha a conexão com o banco de dados se estiver aberta
        if 'conn' in locals():
            conn.close()

            
def obter_ultima_pk_tabela_estrutura():
    query_ultima_pk_tabela_estrutura = f"""SELECT TOP 1 R_E_C_N_O_ FROM {database}.dbo.SG1010 ORDER BY R_E_C_N_O_ DESC;"""
    
    try:
        # Uso do Context Manager para garantir o fechamento adequado da conexão
        with pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}') as conn:
            cursor = conn.cursor()
            cursor.execute(query_ultima_pk_tabela_estrutura)
            resultado_ultima_pk_tabela_estrutura = cursor.fetchone()
            
            ctypes.windll.user32.MessageBoxW(0, f"{resultado_ultima_pk_tabela_estrutura[0]}", "Valor pk", 0)

            return resultado_ultima_pk_tabela_estrutura[0]

    except Exception as ex:
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou consulta. Erro: {str(ex)}", "Erro ultima pk", 0)
        return None
    
    
def obter_revisao_inicial_codigo_pai(codigo_pai):
    query_revisao_inicial = f"""SELECT B1_REVATU FROM {database}.dbo.SB1010 WHERE B1_COD = '{codigo_pai}'"""
    
    try:
        # Uso do Context Manager para garantir o fechamento adequado da conexão
        with pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}') as conn:
            cursor = conn.cursor()
            cursor.execute(query_revisao_inicial)
            revisao_inicial = cursor.fetchone()
            valor_revisao_inicial = revisao_inicial[0]
            
            if valor_revisao_inicial in ('000', '   '):
                valor_revisao_inicial = '001'
            
            ctypes.windll.user32.MessageBoxW(0, f"{codigo_pai} -> REV. {valor_revisao_inicial}", "last rev value", 0)

            return valor_revisao_inicial

    except Exception as ex:
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou consulta. Erro: {str(ex)}{revisao_inicial[0]}", "Erro Revisão", 0)
        return None
    
    
def obter_unidade_medida_codigo_filho(codigo_filho):
    query_unidade_medida_codigo_filho = f"""SELECT B1_UM FROM {database}.dbo.SB1010 WHERE B1_COD = '{codigo_filho}'"""
    
    try:
        with pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}') as conn:
            cursor = conn.cursor()
            cursor.execute(query_unidade_medida_codigo_filho)
            unidade_medida = cursor.fetchone()
            valor_unidade_medida = unidade_medida[0]
            
            ctypes.windll.user32.MessageBoxW(0, f"{codigo_filho} -> REV. {valor_unidade_medida}", "Unidade Medida", 0)

            return valor_unidade_medida

    except Exception as ex:
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou consulta. Erro: {str(ex)}", "Erro Unidade", 0)
        return None
    
            
def formatar_data_atual():
    # Formato yyyymmdd
    data_atual_formatada = date.today().strftime("%Y%m%d")
    return data_atual_formatada
        
            
def criar_nova_estrutura_totvs(codigo_pai, bom_excel_sem_duplicatas): 
    
    ultima_pk_tabela_estrutura = obter_ultima_pk_tabela_estrutura()
    revisao_inicial = obter_revisao_inicial_codigo_pai(codigo_pai)
    revisao_final = revisao_inicial
    data_atual_formatada = formatar_data_atual()
    
    conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
    
    try:
        
        cursor = conn.cursor()
            
        for index, row in bom_excel_sem_duplicatas.iterrows():
            ultima_pk_tabela_estrutura += 1
            codigo_filho = row.iloc[posicao_coluna_codigo_excel]
            quantidade = row.iloc[posicao_coluna_quantidade_excel]
            unidade_medida = obter_unidade_medida_codigo_filho(codigo_filho)
            
            query_criar_nova_estrutura_totvs = f"""
                INSERT INTO {database}.dbo.SG1010 
                (G1_FILIAL, G1_COD, G1_COMP, G1_TRT, G1_XUM, G1_QUANT, G1_PERDA, G1_INI, G1_FIM, G1_OBSERV, G1_FIXVAR, G1_GROPC, G1_OPC, G1_REVINI, G1_NIV, G1_NIVINV, G1_REVFIM, 
                G1_OK, G1_POTENCI, G1_TIPVEC, G1_VECTOR, G1_VLCOMPE, G1_LOCCONS, G1_USAALT, G1_FANTASM, G1_LISTA, D_E_L_E_T_, R_E_C_N_O_, R_E_C_D_E_L_) 
                VALUES (N'0101', N'{codigo_pai}  ', N'{codigo_filho}  ', N'   ', N'{unidade_medida}', {quantidade}, 0.0, N'{data_atual_formatada}', N'20491231', 
                N'                                             ', N'V', N'   ', N'    ', N'{revisao_inicial}', N'01', N'99', N'{revisao_final}', N'    ', 0.0, N'      ', N'      ', 
                N'N', N'  ', N'1', N' ', N'          ', 
                N' ', {ultima_pk_tabela_estrutura}, 0);
            """
            
            cursor.execute(query_criar_nova_estrutura_totvs)
            
        conn.commit()
        
    except Exception as ex:
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou consulta. Erro: {str(ex)}{unidade_medida}", "Erro Criar Nova Estrutura", 0)
        
    finally:
        cursor.close()
        conn.close()


def alterar_estrutura_existente(bom_excel_sem_duplicatas, resultado_query_consulta_estrutura_totvs):
    # Remove espaços em branco da coluna 'G1_COD'
    resultado_query_consulta_estrutura_totvs['G1_COMP'] = resultado_query_consulta_estrutura_totvs['G1_COMP'].str.strip()

    # Encontra códigos que são iguais entre SQL e Excel
    codigos_em_comum = resultado_query_consulta_estrutura_totvs['G1_COMP'].loc[resultado_query_consulta_estrutura_totvs['G1_COMP'].isin(
        bom_excel_sem_duplicatas.iloc[:, posicao_coluna_codigo_excel])].tolist()
    codigos_adicionados_bom = bom_excel_sem_duplicatas.iloc[:, posicao_coluna_codigo_excel].loc[~bom_excel_sem_duplicatas.iloc[:, posicao_coluna_codigo_excel].isin(
        resultado_query_consulta_estrutura_totvs['G1_COMP'])].tolist()
    codigos_removidos_bom = resultado_query_consulta_estrutura_totvs['G1_COMP'].loc[~resultado_query_consulta_estrutura_totvs['G1_COMP'].isin(
        bom_excel_sem_duplicatas.iloc[:, posicao_coluna_codigo_excel])].tolist()

    # Exibe uma caixa de diálogo com base nos resultados
    if codigos_em_comum:
        ctypes.windll.user32.MessageBoxW(
            0, f"Códigos em comuns: {codigos_em_comum}", "ITENS EM COMUM", 1)

    if codigos_adicionados_bom:
        ctypes.windll.user32.MessageBoxW(
            0, f"Itens adicionados: {codigos_adicionados_bom}", "ITENS ADICIONADOS", 1)

    if codigos_removidos_bom:
        ctypes.windll.user32.MessageBoxW(
            0, f"Itens removidos: {codigos_removidos_bom}", "ITENS REMOVIDOS", 1)

nome_desenho = ler_variavel_ambiente_codigo_desenho()
excel_file_path = obter_caminho_arquivo_excel(nome_desenho)

bom_excel_sem_duplicatas = validacao_de_dados_bom(excel_file_path)
resultado_query_consulta_estrutura_totvs = verificar_se_existe_estrutura_totvs(nome_desenho)

if resultado_query_consulta_estrutura_totvs.empty:
    criar_nova_estrutura_totvs(nome_desenho, bom_excel_sem_duplicatas)
#else:
    #alterar_estrutura_existente(bom_excel_sem_duplicatas, resultado_query_consulta_estrutura_totvs)

delete_file_if_exists(excel_file_path)
