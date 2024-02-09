import pyodbc
import pandas as pd
import ctypes
import os
import re
from datetime import date
import tkinter as tk
from tkinter import messagebox
from sqlalchemy import create_engine

# Parâmetros de conexão com o banco de dados SQL Server
server = 'SVRERP,1433'
database = 'PROTHEUS12_R27' # PROTHEUS12_R27 (base de produção) PROTHEUS1233_HML (base de desenvolvimento/teste)
username = 'coognicao'
password = '0705@Abc'
driver = '{ODBC Driver 17 for SQL Server}'

titulo_janela = "CADASTRO DE ESTRUTURA - TOTVS®"

# Arrays para armazenar os códigos
codigos_adicionados_bom = []  # ITENS ADICIONADOS
codigos_removidos_bom = []  # ITENS REMOVIDOS
codigos_em_comum = []  # ITENS EM COMUM

indice_coluna_codigo_excel = 1
indice_coluna_descricao_excel = 2
indice_coluna_quantidade_excel = 3
indice_coluna_peso_excel = 6

formatos_codigo = [
        r'^(C|M)\-\d{3}\-\d{3}\-\d{3}$',
        r'^(E\d{4}\-\d{3}\-\d{3})$',
        r'^(E\d{4}\-\d{3}\-A\d{2})$',
        r'^(E\d{12})$',
    ]

def validar_formato_codigo_pai(codigo_pai):  
    codigo_pai_validado = any(re.match(formato, str(codigo_pai)) for formato in formatos_codigo)
    
    if not codigo_pai_validado:
        exibir_mensagem(titulo_janela, "Este desenho está com o código fora do formato padrão ENAPLIC.\n\nCÓDIGO {codigo_pai}\n\nCorrija e tente novamente!\n\nツ", "info")

    return codigo_pai_validado
    

def validar_formato_codigos_filho(df_excel):
    validacao_codigos = df_excel.iloc[:, indice_coluna_codigo_excel].apply(
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
        exibir_mensagem(titulo_janela, f"PRODUTOS REPETIDOS NA BOM.\n\nOs códigos são iguais com descrições diferentes:\n\n{codigos_repetidos.tolist()}\n\nCorrija-os ou exclue da tabela e tente novamente!\n\nツ", "warning")
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
            mensagem = f"CÓDIGOS-FILHO SEM CADASTRO NO TOTVS:\n\n{', '.join(codigos_sem_cadastro)}\n\nEfetue o cadastro e tente novamente!\n\nツ"
            exibir_mensagem(titulo_janela, mensagem, "info")
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


def verificar_se_existe_estrutura_codigos_filho(codigos_filho):            
    try:
        conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
        cursor = conn.cursor()

        codigos_sem_estrutura = []
        
        for codigo_produto in codigos_filho:
            
            query_consulta_tipo_produto = f"""SELECT B1_TIPO FROM {database}.dbo.SB1010 WHERE B1_COD = '{codigo_produto}' AND B1_TIPO IN ('PI','PA');"""
            
            cursor.execute(query_consulta_tipo_produto)
            resultado_tipo_produto = cursor.fetchone()
            
            if resultado_tipo_produto:
            
                query_consulta_estrutura_totvs = f"""SELECT *
                    FROM {database}.dbo.SG1010
                    WHERE G1_COD = '{codigo_produto}' 
                    AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*'
                    AND G1_REVFIM = (SELECT MAX(G1_REVFIM) FROM {database}.dbo.SG1010 WHERE G1_COD = '{codigo_produto}'AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*');
                """
                
                cursor.execute(query_consulta_estrutura_totvs)
                resultado = cursor.fetchone()

                if not resultado:
                    codigos_sem_estrutura.append(codigo_produto)

        if codigos_sem_estrutura:
            mensagem = f"CÓDIGOS-FILHO SEM ESTRUTURA NO TOTVS:\n\n{', '.join(codigos_sem_estrutura)}\n\nEfetue o cadastro da estrutura de cada um deles e tente novamente!\n\nツ"
            exibir_mensagem(titulo_janela, mensagem, "info")
            return False
        else:
            return True

    except Exception as ex:
        # Exibe uma caixa de diálogo se a conexão ou a consulta falhar
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}", "Erro ao consultar o cadastro de estrutura dos itens filho", 16 | 0)
        
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


def validacao_quantidades(df_excel):
    return df_excel.iloc[:, indice_coluna_quantidade_excel].notna() & (df_excel.iloc[:, indice_coluna_quantidade_excel] != '') & (pd.to_numeric(df_excel.iloc[:, indice_coluna_quantidade_excel], errors='coerce') > 0)   


def validacao_pesos(df_excel):
    return df_excel.iloc[:, indice_coluna_peso_excel].notna() & ((df_excel.iloc[:, indice_coluna_peso_excel] == 0) | (pd.to_numeric(df_excel.iloc[:, indice_coluna_peso_excel], errors='coerce') > 0))


def validacao_de_dados_bom(excel_file_path):
    # Carrega a planilha do Excel em um DataFrame
    df_excel = pd.read_excel(excel_file_path, sheet_name='Planilha1', header=None)

    # Exclui a última linha do DataFrame
    df_excel = df_excel.drop(df_excel.index[-1])
    
    codigo_filho_diferente_codigo_pai = verificar_codigo_filho_diferente_codigo_pai(nome_desenho, df_excel)

    validar_codigos = validar_formato_codigos_filho(df_excel)

    validar_quantidades = validacao_quantidades(df_excel)
    
    validar_descricoes = validar_descricao(df_excel.iloc[:, indice_coluna_descricao_excel])

    validar_pesos = validacao_pesos(df_excel)
    
    if not codigo_filho_diferente_codigo_pai.all():
        exibir_mensagem(titulo_janela, "EXISTE CÓDIGO-FILHO NA BOM IGUAL AO CÓDIGO PAI\n\nPor favor, corrija o código e tente novamente!\n\nツ", "info")

    if not validar_codigos.all():
        exibir_mensagem(titulo_janela, "CÓDIGO-FILHO FORA DO FORMATO PADRÃO ENAPLIC\n\nPor favor, corrija o código e tente novamente!\n\nツ", "info")

    if not validar_descricoes.all():
        exibir_mensagem(titulo_janela, "DESCRIÇÃO INVÁLIDA ENCONTRADA\n\nAs descrições não podem ser nulas, vazias ou conter apenas espaços em branco.\nPor favor, corrija a descrição e tente novamente!\n\nツ", "info")

    if not validar_quantidades.all():
        exibir_mensagem(titulo_janela, "QUANTIDADE INVÁLIDA ENCONTRADA\n\nAs quantidades devem ser números, não nulas, sem espaços em branco e maiores que zero.\nPor favor, corrija a quantidade e tente novamente!\n\nツ", "info")

    if not validar_pesos.all():
        exibir_mensagem(titulo_janela, "PESO INVÁLIDO ENCONTRADO\n\nOs pesos devem ser números, não nulos, sem espaços em branco e maiores ou iguais à zero.\nPor favor, corrija-os e tente novamente!\n\nツ", "info")
    
    if validar_codigos.all() and validar_descricoes.all() and validar_quantidades.all() and codigo_filho_diferente_codigo_pai.all() and validar_pesos.all():

        bom_excel_sem_duplicatas = remover_linhas_duplicadas_e_consolidar_quantidade(df_excel)
        bom_excel_sem_duplicatas.iloc[:, indice_coluna_codigo_excel] = bom_excel_sem_duplicatas.iloc[:, indice_coluna_codigo_excel].str.strip()
        
        existe_codigo_filho_repetido = verificar_codigo_repetido(bom_excel_sem_duplicatas)        
        codigos_filho_tem_cadastro = verificar_cadastro_codigo_filho(bom_excel_sem_duplicatas.iloc[:, indice_coluna_codigo_excel].tolist())
        codigos_filho_tem_estrutura = verificar_se_existe_estrutura_codigos_filho(bom_excel_sem_duplicatas.iloc[:, indice_coluna_codigo_excel].tolist())
        
        if not existe_codigo_filho_repetido and codigos_filho_tem_cadastro and codigos_filho_tem_estrutura:
            return bom_excel_sem_duplicatas
        else:
            return None


def atualizar_campo_revisao_do_codigo_pai(codigo_pai, numero_revisao):
    query_atualizar_campo_revisao = f"""UPDATE {database}.dbo.SB1010 SET B1_REVATU = N'{numero_revisao}' WHERE B1_COD = N'{codigo_pai}';"""
    try:
        # Uso do Context Manager para garantir o fechamento adequado da conexão
        with pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}') as conn:
            cursor = conn.cursor()
            cursor.execute(query_atualizar_campo_revisao)
            return True

    except Exception as ex:
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}", "Erro ao atualizar o campo revisão do código pai", 16 | 0)
        return False
    

def verificar_se_existe_estrutura_codigo_pai(codigo_pai):
    try:
        # Tente estabelecer a conexão com o banco de dados
        conn_str = f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        engine = create_engine(f'mssql+pyodbc:///?odbc_connect={conn_str}')

        query_consulta_estrutura_totvs = f"""
            SELECT * FROM {database}.dbo.SG1010
            WHERE G1_COD = '{codigo_pai}' AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*'
                AND G1_REVFIM = (
                    SELECT MAX(G1_REVFIM)
                    FROM {database}.dbo.SG1010
                    WHERE G1_COD = '{codigo_pai}' AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*'
                );
        """

        # Executa a query SELECT e obtém os resultados em um DataFrame
        resultado_query_consulta_estrutura_totvs = pd.read_sql(query_consulta_estrutura_totvs, engine)

        return resultado_query_consulta_estrutura_totvs

    except Exception as ex:
        # Exibe uma caixa de diálogo se a conexão ou a consulta falhar
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}", "Erro ao verificar se existe estrutura no TOTVS", 16 | 0)

    finally:
        # Fecha a conexão com o banco de dados se estiver aberta
        if 'engine' in locals():
            engine.dispose()

            
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
    
    
def obter_revisao_codigo_pai(codigo_pai, primeiro_cadastro):
    query_revisao_inicial = f"""SELECT B1_REVATU FROM {database}.dbo.SB1010 WHERE B1_COD = '{codigo_pai}'"""   
    try:
        # Uso do Context Manager para garantir o fechamento adequado da conexão
        with pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}') as conn:
            cursor = conn.cursor()
            cursor.execute(query_revisao_inicial)
            revisao_inicial = cursor.fetchone()
            valor_revisao_inicial = revisao_inicial[0]
            
            if primeiro_cadastro and valor_revisao_inicial in ('000', '001', '   '):
                valor_revisao_inicial = '001'
            elif not primeiro_cadastro and valor_revisao_inicial not in ('000', '   '):
                valor_revisao_inicial = int(valor_revisao_inicial) + 1
                if valor_revisao_inicial <= 9:
                    valor_revisao_inicial = "00" + str(valor_revisao_inicial)  
                else:
                    valor_revisao_inicial = "0" + str(valor_revisao_inicial)            

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
                exibir_mensagem(titulo_janela, f"O cadastro do item pai não foi encontrado!\n\nEfetue o cadastro do código {codigo_pai} e, em seguida, tente novamente.\n\nツ", "warning") 
                return False
        
    except Exception as ex:
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}", "Erro ao consultar cadastro do código-pai", 16 | 0)
        return None
    
            
def criar_nova_estrutura_totvs(codigo_pai, bom_excel_sem_duplicatas): 
    primeiro_cadastro = True
    ultima_pk_tabela_estrutura = obter_ultima_pk_tabela_estrutura()
    revisao_inicial = obter_revisao_codigo_pai(codigo_pai, primeiro_cadastro)
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
                
            quantidade_formatada = "{:.2f}".format(float(quantidade))
            
            query_criar_nova_estrutura_totvs = f"""
                INSERT INTO {database}.dbo.SG1010 
                (G1_FILIAL, G1_COD, G1_COMP, G1_TRT, G1_XUM, G1_QUANT, G1_PERDA, G1_INI, G1_FIM, G1_OBSERV, G1_FIXVAR, G1_GROPC, G1_OPC, G1_REVINI, G1_NIV, G1_NIVINV, G1_REVFIM, 
                G1_OK, G1_POTENCI, G1_TIPVEC, G1_VECTOR, G1_VLCOMPE, G1_LOCCONS, G1_USAALT, G1_FANTASM, G1_LISTA, D_E_L_E_T_, R_E_C_N_O_, R_E_C_D_E_L_) 
                VALUES (N'0101', N'{codigo_pai}  ', N'{codigo_filho}  ', N'   ', N'{unidade_medida}', {quantidade_formatada}, 0.0, N'{data_atual_formatada}', N'20491231', 
                N'                                             ', N'V', N'   ', N'    ', N'{revisao_inicial}', N'01', N'99', N'{revisao_inicial}', N'    ', 0.0, N'      ', N'      ', 
                N'N', N'  ', N'1', N' ', N'          ', 
                N' ', {ultima_pk_tabela_estrutura}, 0);
            """
            
            cursor.execute(query_criar_nova_estrutura_totvs)
            
        conn.commit()
        
        exibir_mensagem(titulo_janela, f"ESTRUTURA CADASTRADA COM SUCESSO!\n\n{codigo_pai}\n\nEngenharia ENAPLIC®\n\n( ͡° ͜ʖ ͡°)", "info")
        return True, revisao_inicial
        
    except Exception as ex:
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou consulta. Erro: {str(ex)} - PK-{ultima_pk_tabela_estrutura} - {codigo_pai} - {codigo_filho} - {quantidade} - {unidade_medida}", "Erro ao Criar Nova Estrutura", 16 | 0)
        return False
        
    finally:
        cursor.close()
        conn.close()


def janela_mensagem_alterar_estrutura(codigo_pai):
    user_choice = messagebox.askquestion(
        titulo_janela,
        f"ESTRUTURA EXISTENTE\n\nJá existe uma estrutura cadastrada no TOTVS para este produto!\n\n{codigo_pai}\n\nDeseja realizar a alteração da estrutura?"
    )

    if user_choice == "yes":
        return True
    else:
        return False
    
    
def atualizar_itens_estrutura_totvs(codigo_pai, dataframe_codigos_em_comum):
    conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
    
    try:
        
        cursor = conn.cursor()
            
        for index, row in dataframe_codigos_em_comum.iterrows():
            codigo_filho = row.iloc[indice_coluna_codigo_excel]
            quantidade = row.iloc[indice_coluna_quantidade_excel]
            unidade_medida = obter_unidade_medida_codigo_filho(codigo_filho)
            
            if unidade_medida == 'KG':
                quantidade = row.iloc[indice_coluna_peso_excel]
                
            quantidade_formatada = "{:.2f}".format(float(quantidade))
            
            query_alterar_quantidade_estrutura = f"""UPDATE {database}.dbo.SG1010 SET G1_QUANT = {quantidade_formatada} WHERE G1_COD = '{codigo_pai}' AND G1_COMP = '{codigo_filho}'
                AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*'
                AND G1_REVFIM = (SELECT MAX(G1_REVFIM) FROM {database}.dbo.SG1010 WHERE G1_COD = '{codigo_pai}' AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*');
            """  
            
            cursor.execute(query_alterar_quantidade_estrutura)
            
        conn.commit()
        
    except Exception as ex:
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou consulta. Erro: {str(ex)}", "Erro ao atualizar itens já existentes da estrutura", 16 | 0)
        
    finally:
        cursor.close()
        conn.close()


def inserir_itens_estrutura_totvs(codigo_pai, dataframe_codigos_adicionados_bom, revisao_atualizada_estrutura):
    ultima_pk_tabela_estrutura = obter_ultima_pk_tabela_estrutura()
    data_atual_formatada = formatar_data_atual()
    revisao_inicial = revisao_atualizada_estrutura
    conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
    
    try:
        
        cursor = conn.cursor()
            
        for index, row in dataframe_codigos_adicionados_bom.iterrows():
            ultima_pk_tabela_estrutura += 1
            codigo_filho = row.iloc[indice_coluna_codigo_excel]
            quantidade = row.iloc[indice_coluna_quantidade_excel]
            unidade_medida = obter_unidade_medida_codigo_filho(codigo_filho)
            
            if unidade_medida == 'KG':
                quantidade = row.iloc[indice_coluna_peso_excel]
                
            quantidade_formatada = "{:.2f}".format(float(quantidade))
            
            query_criar_nova_estrutura_totvs = f"""
                INSERT INTO {database}.dbo.SG1010 
                (G1_FILIAL, G1_COD, G1_COMP, G1_TRT, G1_XUM, G1_QUANT, G1_PERDA, G1_INI, G1_FIM, G1_OBSERV, G1_FIXVAR, G1_GROPC, G1_OPC, G1_REVINI, G1_NIV, G1_NIVINV, G1_REVFIM, 
                G1_OK, G1_POTENCI, G1_TIPVEC, G1_VECTOR, G1_VLCOMPE, G1_LOCCONS, G1_USAALT, G1_FANTASM, G1_LISTA, D_E_L_E_T_, R_E_C_N_O_, R_E_C_D_E_L_) 
                VALUES (N'0101', N'{codigo_pai}  ', N'{codigo_filho}  ', N'   ', N'{unidade_medida}', {quantidade_formatada}, 0.0, N'{data_atual_formatada}', N'20491231', 
                N'                                             ', N'V', N'   ', N'    ', N'{revisao_inicial}', N'01', N'99', N'{revisao_atualizada_estrutura}', N'    ', 0.0, N'      ', N'      ', 
                N'N', N'  ', N'1', N' ', N'          ', 
                N' ', {ultima_pk_tabela_estrutura}, 0);
            """
            
            cursor.execute(query_criar_nova_estrutura_totvs)
            
        conn.commit()

        return True
    
    except Exception as ex:
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou consulta. Erro: {str(ex)}", "Erro ao inserir novos itens na estrutura", 16 | 0)
        return False
        
    finally:
        cursor.close()
        conn.close()


def remover_itens_estrutura_totvs(codigo_pai, codigos_removidos_bom_df, revisao_anterior):
    conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
    
    try:
        
        cursor = conn.cursor()
            
        for index, row in codigos_removidos_bom_df.iterrows():
            codigo_filho = row.iloc[2]
            
            query_remover_itens_estrutura_totvs = f"""
            UPDATE {database}.dbo.SG1010
            SET
                D_E_L_E_T_ = N'*',
                R_E_C_D_E_L_ = R_E_C_N_O_
            WHERE
                G1_COD = '{codigo_pai}' AND G1_COMP = '{codigo_filho}'
                AND G1_REVFIM = N'{revisao_anterior}'
                AND G1_REVFIM <> 'ZZZ'
                AND D_E_L_E_T_ <> '*';
            """  
            
            cursor.execute(query_remover_itens_estrutura_totvs)
            
        conn.commit()
        
        return True
        
    except Exception as ex:
        return False
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou consulta. Erro: {str(ex)}", "Erro ao remover item da estrutura", 16 | 0)
        
    finally:
        cursor.close()
        conn.close()


def resultado_comparacao():
    if codigos_em_comum:  
        ctypes.windll.user32.MessageBoxW(
            0, f"Códigos em comuns: {codigos_em_comum}", "ITENS EM COMUM", 1)
        
    if codigos_adicionados_bom:         
        ctypes.windll.user32.MessageBoxW(
            0, f"Itens adicionados: {codigos_adicionados_bom}", "ITENS ADICIONADOS", 1)

    if codigos_removidos_bom:  
        ctypes.windll.user32.MessageBoxW(
            0, f"Itens removidos: {codigos_removidos_bom}", "ITENS REMOVIDOS", 1)


def comparar_bom_com_totvs(bom_excel_sem_duplicatas, resultado_query_consulta_estrutura_totvs):   
    resultado_query_consulta_estrutura_totvs['G1_COMP'] = resultado_query_consulta_estrutura_totvs['G1_COMP'].str.strip()

    # Códigos em comum
    codigos_em_comum_df = bom_excel_sem_duplicatas[bom_excel_sem_duplicatas.iloc[:, indice_coluna_codigo_excel].isin(
        resultado_query_consulta_estrutura_totvs['G1_COMP'])].copy()

    # Códigos adicionados no BOM
    codigos_adicionados_bom_df = bom_excel_sem_duplicatas[~bom_excel_sem_duplicatas.iloc[:, indice_coluna_codigo_excel].isin(
        resultado_query_consulta_estrutura_totvs['G1_COMP'])].copy()

    # Códigos removidos no BOM
    codigos_removidos_bom_df = resultado_query_consulta_estrutura_totvs[~resultado_query_consulta_estrutura_totvs['G1_COMP'].isin(
        bom_excel_sem_duplicatas.iloc[:, indice_coluna_codigo_excel])].copy()

    return codigos_em_comum_df, codigos_adicionados_bom_df, codigos_removidos_bom_df
    
    #resultado_comparacao()
    

def atualizar_campo_revfim_codigos_existentes(codigo_pai, revisao_anterior, revisao_atualizada):
    conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}') 
    try:
        cursor = conn.cursor()
         
        for index, row in codigos_em_comum_df.iterrows():
            codigo_filho = row.iloc[indice_coluna_codigo_excel]
            
            query_atualizar_campo_revfim_estrutura = f"""UPDATE {database}.dbo.SG1010 SET G1_REVFIM = N'{revisao_atualizada}' WHERE G1_COD = '{codigo_pai}' AND G1_COMP = '{codigo_filho}'
               AND G1_REVFIM = N'{revisao_anterior}' AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*'
            """  
                
            cursor.execute(query_atualizar_campo_revfim_estrutura)
            
        conn.commit()
        
        return True
        
    except Exception as ex:
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou consulta. Erro: {str(ex)}", "Erro ao atualizar campo revfim dos itens já existentes da estrutura", 16 | 0)
        return False
        
    finally:
        cursor.close()
        conn.close()
        

def calculo_revisao_anterior(revisao_atualizada):
    revisao_atualizada = int(revisao_atualizada) - 1
    revisao_anterior = str(revisao_atualizada).zfill(3)
    return revisao_anterior


def exibir_mensagem(title, message, icon_type):
    root = tk.Tk()
    root.withdraw()
    root.lift()  # Garante que a janela esteja na frente
    root.title(title)
    #root.iconbitmap(default='info.ico')  # Substitua 'info.ico' pelo caminho do ícone desejado

    if icon_type == 'info':
        messagebox.showinfo(title, message)
    elif icon_type == 'warning':
        messagebox.showwarning(title, message)
    elif icon_type == 'error':
        messagebox.showerror(title, message)

    root.destroy()


nome_desenho = ler_variavel_ambiente_codigo_desenho()
excel_file_path = obter_caminho_arquivo_excel(nome_desenho)
formato_codigo_pai_correto = validar_formato_codigo_pai(nome_desenho)
revisao_atualizada = None
nova_estrutura_cadastrada = False

if formato_codigo_pai_correto:
    existe_cadastro_codigo_pai = verificar_cadastro_codigo_pai(nome_desenho)

if formato_codigo_pai_correto and existe_cadastro_codigo_pai:
    bom_excel_sem_duplicatas = validacao_de_dados_bom(excel_file_path)
    resultado_estrutura_codigo_pai = verificar_se_existe_estrutura_codigo_pai(nome_desenho)
    
    excluir_arquivo_excel_bom(excel_file_path)

    if not bom_excel_sem_duplicatas.empty and resultado_estrutura_codigo_pai.empty:
        nova_estrutura_cadastrada, revisao_atualizada = criar_nova_estrutura_totvs(nome_desenho, bom_excel_sem_duplicatas)
        atualizar_campo_revisao_do_codigo_pai(nome_desenho, revisao_atualizada)
            
    if not bom_excel_sem_duplicatas.empty and not resultado_estrutura_codigo_pai.empty:
        usuario_quer_alterar = janela_mensagem_alterar_estrutura(nome_desenho)

        if usuario_quer_alterar:
            resultado = comparar_bom_com_totvs(bom_excel_sem_duplicatas, resultado_estrutura_codigo_pai)
            codigos_em_comum_df, codigos_adicionados_bom_df, codigos_removidos_bom_df = resultado
            
            if not codigos_em_comum_df.empty:
                atualizar_itens_estrutura_totvs(nome_desenho, codigos_em_comum_df)
                
            if not codigos_adicionados_bom_df.empty or not codigos_removidos_bom_df.empty:
                primeiro_cadastro = False
                revisao_atualizada = obter_revisao_codigo_pai(nome_desenho, primeiro_cadastro)
                itens_adicionados_sucesso = False
                itens_removidos_sucesso = False
                revisao_anterior = calculo_revisao_anterior(revisao_atualizada)
                
                if not codigos_adicionados_bom_df.empty:         
                    itens_adicionados_sucesso = inserir_itens_estrutura_totvs(nome_desenho, codigos_adicionados_bom_df, revisao_atualizada)

                if not codigos_removidos_bom_df.empty:  
                    itens_removidos_sucesso = remover_itens_estrutura_totvs(nome_desenho, codigos_removidos_bom_df, revisao_anterior)
                    
                if itens_adicionados_sucesso or itens_removidos_sucesso:
                    atualizar_campo_revfim_codigos_existentes(nome_desenho, revisao_anterior, revisao_atualizada)
                    atualizar_campo_revisao_do_codigo_pai(nome_desenho, revisao_atualizada)                    
                    exibir_mensagem(titulo_janela, f"ATUALIZAÇÃO DA ESTRUTURA REALIZADA COM SUCESSO!\n\n{nome_desenho}\n\nEngenharia ENAPLIC®\n\n( ͡° ͜ʖ ͡°)", "info")
            else:
                exibir_mensagem(titulo_janela,f"QUANTIDADES ATUALIZADAS COM SUCESSO!\n\nNenhum item foi ADICIONADO e/ou REMOVIDO.\n\n{nome_desenho}\n\nEngenharia ENAPLIC®\n\n( ͡° ͜ʖ ͡°)","info")
    elif not nova_estrutura_cadastrada:
        exibir_mensagem(titulo_janela,f"OPS!\n\nA BOM está vazia!\n\nPor gentileza, preencha adequadamente a BOM e tente novamente!\n\n{nome_desenho}\n\nEngenharia ENAPLIC®\n\nツ","warning")
else:
    excluir_arquivo_excel_bom(excel_file_path)