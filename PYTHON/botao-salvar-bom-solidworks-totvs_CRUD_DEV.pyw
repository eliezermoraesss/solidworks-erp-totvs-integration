import pyodbc
import pandas as pd
import ctypes
import os
import re
from datetime import date
import tkinter as tk
from tkinter import messagebox

# Parâmetros de conexão com o banco de dados SQL Server
server = 'SVRERP,1433'
database = 'PROTHEUS1233_HML' # PROTHEUS12_R27 (base de produção) PROTHEUS1233_HML (base de desenvolvimento/teste)
username = 'coognicao'
password = '0705@Abc'
driver = '{ODBC Driver 17 for SQL Server}'

# Arrays para armazenar os códigos
codigos_adicionados_bom = []  # ITENS ADICIONADOS
codigos_removidos_bom = []  # ITENS REMOVIDOS
codigos_em_comum = []  # ITENS EM COMUM

indice_coluna_codigo_excel = 1
indice_coluna_descricao_excel = 2
indice_coluna_quantidade_excel = 3
indice_coluna_peso_excel = 6

def validar_formato_codigo_pai(codigo_pai):
    
    formatos_codigo = [
        r'^(C|M)\-\d{3}\-\d{3}\-\d{3}$',
        r'^(E\d{4}\-\d{3}\-\d{3})$',
    ]
    
    codigo_pai_validado = any(re.match(formato, str(codigo_pai)) for formato in formatos_codigo)
    
    if not codigo_pai_validado:
        ctypes.windll.user32.MessageBoxW(0, f"Este desenho está com o código fora do formato padrão ENAPLIC.\n\nCÓDIGO {codigo_pai}\n\nCorrija e tente novamente!", "CADASTRO DE ESTRUTURA - TOTVS®", 64 | 0) 
    
    return codigo_pai_validado
    

def validar_formato_codigos_filho(df_excel, posicao_coluna_codigo):
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


def excluir_arquivo_excel_bom(excel_file_path):
    if os.path.exists(excel_file_path):
        os.remove(excel_file_path)


def verificar_codigo_repetido(df_excel):
    
    codigos_repetidos = df_excel.iloc[:, indice_coluna_codigo_excel][df_excel.iloc[:, indice_coluna_codigo_excel].duplicated()]
    
    # Exibe uma mensagem se houver códigos repetidos
    if not codigos_repetidos.empty:
        ctypes.windll.user32.MessageBoxW(
            0, f"Produtos repetidos na BOM.\n\nOs códigos são iguais com descrições diferentes:\n\n{codigos_repetidos.tolist()}\n\nCorrija-os ou exclue da tabela e tente novamente!", "CADASTRO DE ESTRUTURA - TOTVS®", 48 | 0)
        return True
    else:
        return False


def verificar_cadastro_codigo_filho(codigos_filho):            
    try:
        conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
        cursor = conn.cursor()

        codigos_sem_cadastro = []
        
        for codigo_produto in codigos_filho:
            query_consulta_produto = f"""
            SELECT B1_COD FROM {database}.dbo.SB1010 WHERE B1_COD = '{codigo_produto}';
            """
            
            cursor.execute(query_consulta_produto)
            resultado = cursor.fetchone()

            if not resultado:
                codigos_sem_cadastro.append(codigo_produto)

        if codigos_sem_cadastro:
            mensagem = f"Códigos-filho sem cadastro no TOTVS:\n\n{', '.join(codigos_sem_cadastro)}\n\nEfetue o cadastro e tente novamente!"
            ctypes.windll.user32.MessageBoxW(0, mensagem, "CADASTRO DE ESTRUTURA - TOTVS®", 64 | 0)
            return False
        else:
            return True

    except Exception as ex:
        # Exibe uma caixa de diálogo se a conexão ou a consulta falhar
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}", "Erro ao consultar o cadastro de produtos", 16 | 0)
        
    finally:
        # Fecha a conexão com o banco de dados se estiver aberta
        if 'conn' in locals():
            conn.close()
        

def remover_linhas_duplicadas_e_consolidar_quantidade(df_excel):
    # Agrupa o DataFrame pela combinação única de código e descrição
    grouped = df_excel.groupby([indice_coluna_codigo_excel, indice_coluna_descricao_excel])

    # Inicializa um novo DataFrame para armazenar o resultado
    df_sem_duplicatas = pd.DataFrame(columns=df_excel.columns)

    # Itera sobre os grupos consolidando as quantidades
    for _, group in grouped:
        codigo = group[indice_coluna_codigo_excel].iloc[0]
        descricao = group[indice_coluna_descricao_excel].iloc[0]
        quantidade_consolidada = group[indice_coluna_quantidade_excel].sum()
        peso_consolidado = group[indice_coluna_peso_excel].sum()

        # Adiciona uma linha ao DataFrame sem duplicatas
        df_sem_duplicatas = pd.concat([df_sem_duplicatas, group.head(1)])
        df_sem_duplicatas.loc[df_sem_duplicatas.index[-1], indice_coluna_quantidade_excel] = quantidade_consolidada
        df_sem_duplicatas.loc[df_sem_duplicatas.index[-1], indice_coluna_peso_excel] = peso_consolidado

    return df_sem_duplicatas


def validar_descricao(descricoes):
    return descricoes.notna() & (descricoes != '') & (descricoes.astype(str).str.strip() != '')


def verificar_codigo_filho_diferente_codigo_pai(nome_desenho, df_excel):
    codigo_filho_diferente_codigo_pai = df_excel.iloc[:, indice_coluna_codigo_excel] != f"{nome_desenho}"
    return codigo_filho_diferente_codigo_pai

def validacao_de_dados_bom(excel_file_path):

    # Carrega a planilha do Excel em um DataFrame
    df_excel = pd.read_excel(excel_file_path, sheet_name='Planilha1', header=None)

    # Exclui a última linha do DataFrame
    df_excel = df_excel.drop(df_excel.index[-1])
    
    codigo_filho_diferente_codigo_pai = verificar_codigo_filho_diferente_codigo_pai(nome_desenho, df_excel)

    validar_codigos = validar_formato_codigos_filho(df_excel, indice_coluna_codigo_excel)

    validar_quantidades = df_excel.iloc[:, indice_coluna_quantidade_excel].notna() & (df_excel.iloc[:, indice_coluna_quantidade_excel] != '') & (pd.to_numeric(df_excel.iloc[:, indice_coluna_quantidade_excel], errors='coerce') > 0)
    
    validar_descricoes = validar_descricao(df_excel.iloc[:, indice_coluna_descricao_excel])

    if not codigo_filho_diferente_codigo_pai.all():
        ctypes.windll.user32.MessageBoxW(
            0, "EXISTE CÓDIGO-FILHO NA BOM IGUAL AO CÓDIGO PAI\n\nPor favor, corrija o código e tente novamente!", "CADASTRO DE ESTRUTURA - TOTVS®", 48 | 0)
        
    if not validar_codigos.all():
        ctypes.windll.user32.MessageBoxW(
            0, "CÓDIGO-FILHO FORA DO FORMATO PADRÃO ENAPLIC\n\nPor favor, corrija o código e tente novamente!", "CADASTRO DE ESTRUTURA - TOTVS®", 48 | 0)
        
    if not validar_descricoes.all():
        ctypes.windll.user32.MessageBoxW(
            0, "DESCRIÇÃO INVÁLIDA ENCONTRADA\n\nAs descrições não podem ser nulas, vazias ou conter apenas espaços em branco.\nPor favor, corrija o código e tente novamente!", "CADASTRO DE ESTRUTURA - TOTVS®", 48 | 0)

    if not validar_quantidades.all():
        ctypes.windll.user32.MessageBoxW(
            0, "QUANTIDADE INVÁLIDA ENCONTRADA\n\nAs quantidades devem ser números, não nulas, sem espaços em branco e maiores que zero.\nPor favor, corrija o código e tente novamente!", "CADASTRO DE ESTRUTURA - TOTVS®", 48 | 0)

    if validar_codigos.all() and validar_descricoes.all() and validar_quantidades.all() and codigo_filho_diferente_codigo_pai.all():

        bom_excel_sem_duplicatas = remover_linhas_duplicadas_e_consolidar_quantidade(df_excel)
        bom_excel_sem_duplicatas.iloc[:, indice_coluna_codigo_excel] = bom_excel_sem_duplicatas.iloc[:, indice_coluna_codigo_excel].str.strip()
        
        existe_codigo_filho_repetido = verificar_codigo_repetido(bom_excel_sem_duplicatas)        
        codigos_filho_tem_cadastro = verificar_cadastro_codigo_filho(bom_excel_sem_duplicatas.iloc[:, indice_coluna_codigo_excel].tolist())    
        
        if not existe_codigo_filho_repetido and codigos_filho_tem_cadastro:
            return bom_excel_sem_duplicatas
        else:
            bom_excel_sem_duplicatas = None
            return bom_excel_sem_duplicatas


def atualizar_campo_revisao_do_codigo_pai(codigo_pai, numero_revisao):
    query_atualizar_campo_revisao = f"""UPDATE {database}.dbo.SB1010 SET B1_REVATU = '{numero_revisao}' WHERE B1_COD = '{codigo_pai}';"""
    try:
        # Uso do Context Manager para garantir o fechamento adequado da conexão
        with pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}') as conn:
            cursor = conn.cursor()
            cursor.execute(query_atualizar_campo_revisao)
            return True

    except Exception as ex:
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}", "Erro ao atualizar o campo revisão do código pai", 16 | 0)
        return False
    

def verificar_se_existe_estrutura_totvs(codigo_pai):

    query_consulta_estrutura_totvs = f"""SELECT * FROM {database}.dbo.SG1010 WHERE G1_COD = '{codigo_pai}'
        AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*'
        AND G1_REVFIM = (SELECT MAX(G1_REVFIM) FROM {database}.dbo.SG1010 WHERE G1_COD = '{codigo_pai}' AND G1_REVFIM <> 'ZZZ');
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
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}", "Erro ao verificar se existe estrutura no TOTVS", 16 | 0)

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
            valor_ultima_pk = resultado_ultima_pk_tabela_estrutura[0]

            return valor_ultima_pk

    except Exception as ex:
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}", "Erro ao obter última PK da tabela estrutura", 16 | 0)
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

            return valor_revisao_inicial

    except Exception as ex:
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}", "Erro ao consultar revisão do código pai na tabela de produtos", 16 | 0)
        return None
    
    
def obter_unidade_medida_codigo_filho(codigo_filho):
    query_unidade_medida_codigo_filho = f"""SELECT B1_UM FROM {database}.dbo.SB1010 WHERE B1_COD = '{codigo_filho}'"""
    try:
        with pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}') as conn:
            cursor = conn.cursor()
            cursor.execute(query_unidade_medida_codigo_filho)
            
            unidade_medida = cursor.fetchone()
            valor_unidade_medida = unidade_medida[0]

            return valor_unidade_medida

    except Exception as ex:
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}", "Erro ao consultar Unidade de Medida de código-filho", 16 | 0)
        return None
    
            
def formatar_data_atual():
    # Formato yyyymmdd
    data_atual_formatada = date.today().strftime("%Y%m%d")
    return data_atual_formatada


def verificar_cadastro_codigo_pai(codigo_pai):
    query_consulta_produto_codigo_pai = f"""SELECT B1_COD FROM {database}.dbo.SB1010 WHERE B1_COD = '{codigo_pai}'"""  
    try:
        with pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}') as conn:
            cursor = conn.cursor()
            cursor.execute(query_consulta_produto_codigo_pai)
                
            resultado = cursor.fetchone()
            
            if resultado:
                return True
            else:
                ctypes.windll.user32.MessageBoxW(0, f"O cadastro do item pai não foi encontrado!\n\nEfetue o cadastro do código {codigo_pai} e, em seguida, tente novamente.", "CADASTRO DE ESTRUTURA - TOTVS®", 48 | 0) 
                return False
        
    except Exception as ex:
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}", "Erro ao consultar cadastro do código-pai", 16 | 0)
        return None
    
            
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
            codigo_filho = row.iloc[indice_coluna_codigo_excel]
            quantidade = row.iloc[indice_coluna_quantidade_excel]
            unidade_medida = obter_unidade_medida_codigo_filho(codigo_filho)
            
            if unidade_medida == 'KG':
                quantidade = row.iloc[indice_coluna_peso_excel]
            
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
        
        ctypes.windll.user32.MessageBoxW(0, f"A ESTRUTURA FOI CADASTRADA COM SUCESSO!\n\n{codigo_pai}\n\nEngenharia ENAPLIC®\n\n:)", "CADASTRO DE ESTRUTURA - TOTVS®", 0x40 | 0x1)
        return revisao_final
        
    except Exception as ex:
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou consulta. Erro: {str(ex)} - PK-{ultima_pk_tabela_estrutura} - {codigo_pai} - {codigo_filho} - {quantidade} - {unidade_medida}", "Erro ao Criar Nova Estrutura", 16 | 0)
        return None
        
    finally:
        cursor.close()
        conn.close()


def ask_user_for_action(codigo_pai):
    user_choice = messagebox.askquestion(
        "CADASTRO DE ESTRUTURA - TOTVS®",
        f"ESTRUTURA EXISTENTE\n\nJá existe uma estrutura cadastrada no TOTVS para este produto!\n\n{codigo_pai}\n\nDeseja realizar a alteração da estrutura?"
    )

    if user_choice == "yes":
        return True
    else:
        return False

def alterar_estrutura_existente(bom_excel_sem_duplicatas, resultado_query_consulta_estrutura_totvs):
    # Remove espaços em branco da coluna 'G1_COD'
    resultado_query_consulta_estrutura_totvs['G1_COMP'] = resultado_query_consulta_estrutura_totvs['G1_COMP'].str.strip()

    # Encontra códigos que são iguais entre SQL e Excel
    codigos_em_comum = resultado_query_consulta_estrutura_totvs['G1_COMP'].loc[resultado_query_consulta_estrutura_totvs['G1_COMP'].isin(
        bom_excel_sem_duplicatas.iloc[:, indice_coluna_codigo_excel])].tolist()
    codigos_adicionados_bom = bom_excel_sem_duplicatas.iloc[:, indice_coluna_codigo_excel].loc[~bom_excel_sem_duplicatas.iloc[:, indice_coluna_codigo_excel].isin(
        resultado_query_consulta_estrutura_totvs['G1_COMP'])].tolist()
    codigos_removidos_bom = resultado_query_consulta_estrutura_totvs['G1_COMP'].loc[~resultado_query_consulta_estrutura_totvs['G1_COMP'].isin(
        bom_excel_sem_duplicatas.iloc[:, indice_coluna_codigo_excel])].tolist()

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
formato_codigo_pai_correto = validar_formato_codigo_pai(nome_desenho)
nao_existe_estrutura_totvs = False

if formato_codigo_pai_correto:
    existe_cadastro_codigo_pai = verificar_cadastro_codigo_pai(nome_desenho)

if formato_codigo_pai_correto and existe_cadastro_codigo_pai:
    bom_excel_sem_duplicatas = validacao_de_dados_bom(excel_file_path)
    resultado_query_consulta_estrutura_totvs = verificar_se_existe_estrutura_totvs(nome_desenho)
    excluir_arquivo_excel_bom(excel_file_path)

    if not bom_excel_sem_duplicatas.empty and resultado_query_consulta_estrutura_totvs.empty:
        revisao_atualizada = criar_nova_estrutura_totvs(nome_desenho, bom_excel_sem_duplicatas)

        if revisao_atualizada != None:
            atualizar_campo_revisao_do_codigo_pai(nome_desenho, revisao_atualizada)
            
    if not bom_excel_sem_duplicatas.empty and not resultado_query_consulta_estrutura_totvs.empty:
        user_wants_to_alter = ask_user_for_action(nome_desenho)

        if user_wants_to_alter:
            alterar_estrutura_existente(bom_excel_sem_duplicatas, resultado_query_consulta_estrutura_totvs)
else:
    excluir_arquivo_excel_bom(excel_file_path)