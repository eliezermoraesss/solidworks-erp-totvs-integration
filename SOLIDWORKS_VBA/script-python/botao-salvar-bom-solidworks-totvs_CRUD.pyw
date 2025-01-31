import pyodbc
import pandas as pd
import ctypes
import os
import re
from datetime import date
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
from sqlalchemy import create_engine
import sys

from sqlalchemy.exc import SQLAlchemyError


def exibir_mensagem(title, message, icon_type):
    tk_window = tk.Tk()
    tk_window.withdraw()
    tk_window.lift()  # Garante que a janela esteja na frente
    tk_window.title(title)
    tk_window.attributes('-topmost', True)

    if icon_type == 'info':
        messagebox.showinfo(title, message)
    elif icon_type == 'warning':
        messagebox.showwarning(title, message)
    elif icon_type == 'error':
        messagebox.showerror(title, message)
    tk_window.destroy()


def calculo_revisao_anterior(revisao_atualizada):
    revisao_atualizada = int(revisao_atualizada) - 1
    revisao_anterior = str(revisao_atualizada).zfill(3)
    return revisao_anterior


def exibir_janela_mensagem_opcao(titulo, mensagem):
    tk_window = tk.Tk()
    tk_window.withdraw()  # Esconde a janela principal
    tk_window.attributes('-topmost', True)  # Garante que a janela estará sempre no topo
    tk_window.lift()  # Traz a janela para frente
    tk_window.focus_force()  # Força o foco na janela

    # Mostrar a mensagem
    user_choice = messagebox.askquestion(
        titulo,
        mensagem,
        parent=tk_window  # Define a janela principal como pai da mensagem
    )

    # Fechar a janela principal após a escolha do usuário
    tk_window.destroy()

    if user_choice == "yes":
        return True
    else:
        return False


def formatar_data_atual():
    # Formato yyyymmdd
    data_atual_formatada = date.today().strftime("%Y%m%d")
    return data_atual_formatada


def extrair_numero(texto):
    numeros = ''
    for char in texto:
        if char.isdigit():
            numeros += char
    return numeros


def extrair_unidade_medida(valor_dimensao):
    padrao = r'[0-9.]+\s*(m³|m²|m)'
    unidade_extraida = re.findall(padrao, valor_dimensao, re.IGNORECASE)
    return unidade_extraida[0]


def excluir_arquivo_excel_bom(excel_file_path):
    if os.path.exists(excel_file_path):
        os.remove(excel_file_path)


def obter_caminho_arquivo_excel(codigo_desenho):
    base_path = os.environ.get('TEMP')
    return os.path.join(base_path, codigo_desenho + '.xlsx')


def ler_variavel_ambiente_codigo_desenho():
    return os.getenv('CODIGO_DESENHO')


def setup_mssql():
    caminho_do_arquivo = (r"\\192.175.175.4\f\INTEGRANTES\ELIEZER\PROJETO SOLIDWORKS "
                          r"TOTVS\libs-python\user-password-mssql\USER_PASSWORD_MSSQL_PROD.txt")
    try:
        with open(caminho_do_arquivo, 'r') as arquivo:
            string_lida = arquivo.read()
            username, password, database, server = string_lida.split(';')
            return username, password, database, server

    except FileNotFoundError as e:
        ctypes.windll.user32.MessageBoxW(0,
                                         f"Erro ao ler credenciais de acesso ao banco de dados MSSQL: {e}"
                                         f"\n\nBase de dados ERP TOTVS PROTHEUS.\n\nPor favor, informe ao "
                                         f"desenvolvedor/TI sobre o erro exibido.\n\nTenha um bom dia! ツ",
                                         "CADASTRO DE ESTRUTURA - TOTVS®", 16 | 0)
        sys.exit()

    except Exception as e:
        ctypes.windll.user32.MessageBoxW(0, f"Ocorreu um erro ao ler o arquivo: {e}", "CADASTRO DE ESTRUTURA - TOTVS®",
                                         16 | 0)
        sys.exit()


def verificar_se_template_bom_esta_correto(dataframe):
    if dataframe.shape[0] >= 1 and dataframe.shape[1] == 10:
        return "montagem"
    elif dataframe.shape[0] >= 1 and dataframe.shape[1] == 9:
        return "peca"
    elif dataframe.shape[0] >= 1 and dataframe.shape[1] == 7:
        return "corte_soldagem"
    elif dataframe.shape[0] == 0:
        raise Exception(f"ATENÇÃO!\n\nBOM VAZIA!\n\nツ\n\nEUREKA®")
    else:
        raise Exception(f"ATENÇÃO!\n\nO TEMPLATE DA BOM FOI ATUALIZADO"
                        f"\n\nPor favor, selecione o template ENAPLIC atual e tente novamente!\n\nツ\n\nEUREKA®")


def validar_descricao(descricoes):
    validar_descricoes = descricoes.notna() & (descricoes != '') & (descricoes.astype(str).str.strip() != '')
    if not validar_descricoes.all():
        raise ValueError("DESCRIÇÃO INVÁLIDA ENCONTRADA\n\nAs descrições não podem ser nulas, vazias ou "
                         "conter apenas espaços em branco.\nPor favor, corrija a descrição e tente novamente!\n\nツ")


class CadastrarBomTOTVS:
    def __init__(self, window):
        # Leitura dos parâmetros de conexão com o banco de dados SQL Server
        self.itens_removidos = None
        self.itens_adicionados = None
        self.itens_em_comum = None
        self.pai_tem_estrutura = None
        self.username, self.password, self.database, self.server = setup_mssql()
        self.driver = '{SQL Server}'

        window.title("Monitor de progresso")
        self.start_time = time.time()

        self.progress = ttk.Progressbar(window, orient="horizontal", length="300", mode="determinate")
        self.progress.pack(pady=20)

        self.status_label = tk.Label(window, text="")
        self.status_label.pack(pady=30)

        self.titulo_janela = "CADASTRO DE ESTRUTURA TOTVS®"

        # Arrays para armazenar os códigos
        self.itens_adicionados_bom = []  # ITENS ADICIONADOS
        self.itens_removidos_bom = []  # ITENS REMOVIDOS
        self.itens_em_comum = []  # ITENS EM COMUM

        self.indice_coluna_codigo_excel = 1
        self.indice_coluna_descricao_excel = 2
        self.indice_coluna_quantidade_excel = 3
        self.indice_coluna_dimensao = 5
        self.indice_coluna_peso_excel = 6
        self.indice_coluna_nome_arquivo = 8

        self.formatos_codigo = [
            r'^(C|M)\-\d{3}\-\d{3}\-\d{3}$',
            r'^(E\d{4}\-\d{3}\-\d{3})$',
            r'^(E\d{4}\-\d{3}\-A\d{2})$',
            r'^(E\d{12})$',
        ]

        self.regex_campo_dimensao = r'^\d*([,.]?\d+)?[mtMT](²|2|³|3)?(\s*\(.*\))?$'

        self.nome_desenho = ler_variavel_ambiente_codigo_desenho()

    def validar_formato_codigo_pai(self, codigo_pai):
        codigo_pai_validado = any(re.match(formato, str(codigo_pai)) for formato in self.formatos_codigo)

        if not codigo_pai_validado:
            raise ValueError(f"Este desenho foi salvo com o código fora do formato padrão ENAPLIC."
                             f"\n\n{codigo_pai}\n\nPor favor renomear e salvar o desenho com o código no formato "
                             f"padrão.\n\nツ")

        return codigo_pai_validado

    def validar_formato_codigos_filho(self, df_excel):
        validacao_codigos = df_excel.iloc[:, self.indice_coluna_codigo_excel].str.strip().apply(
            lambda x: any(re.match(formato, str(x)) for formato in self.formatos_codigo))
        if not validacao_codigos.all():
            raise ValueError("CÓDIGO FILHO FORA DO FORMATO PADRÃO ENAPLIC\n\nPor favor, corrija o código e "
                             "tente novamente!\n\nツ")

    def verificar_codigo_repetido(self, df_excel):

        codigos_repetidos = df_excel.iloc[:, self.indice_coluna_codigo_excel][
            df_excel.iloc[:, self.indice_coluna_codigo_excel].duplicated()]

        # Exibe uma mensagem se houver códigos repetidos
        if not codigos_repetidos.empty:
            raise ValueError(f"PRODUTOS REPETIDOS NA BOM.\n\nOs códigos são iguais com descrições diferentes:"
                             f"\n\n{codigos_repetidos.tolist()}\n\nCorrija-os ou exclue da tabela e tente novamente!"
                             f"\n\nツ")

    def verificar_cadastro_codigo_filho(self, codigos_filho):
        conn = pyodbc.connect(
            f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};'
            f'PWD={self.password}')
        cursor = conn.cursor()
        try:
            codigos_sem_cadastro = []
            for codigo_produto in codigos_filho:
                query_consulta_produto = f"""
                SELECT 
                    B1_COD 
                FROM 
                    {self.database}.dbo.SB1010 
                WHERE 
                    B1_COD = '{codigo_produto.strip()}' 
                AND 
                    D_E_L_E_T_ <> '*';
                """

                cursor.execute(query_consulta_produto)
                resultado = cursor.fetchone()

                if not resultado:
                    codigos_sem_cadastro.append(codigo_produto)

            if codigos_sem_cadastro:
                mensagem = (f"CÓDIGOS-FILHO SEM CADASTRO NO TOTVS:\n\n{', '.join(codigos_sem_cadastro)}"
                            f"\n\nEfetue o cadastro e tente novamente!\n\nツ\n\nEUREKA®")
                raise ValueError(mensagem)

        except pyodbc.Error as sql_ex:
            # Exibe uma caixa de diálogo se a conexão ou a consulta falhar
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(sql_ex)}",
                                             "Erro ao consultar o cadastro de produtos", 16 | 0)
            raise
        except ValueError:
            raise
        except Exception as ex:
            # Trata outros erros inesperados
            ctypes.windll.user32.MessageBoxW(0, f"Erro inesperado: {str(ex)}",
                                             "Erro ao consultar o cadastro de produtos", 16 | 0)
            raise

        finally:
            # Fecha a conexão com o banco de dados se estiver aberta
            if 'conn' in locals():
                conn.close()

    def verificar_se_existe_estrutura_codigos_filho(self, codigos_filho):
        conn = pyodbc.connect(
            f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};'
            f'PWD={self.password}')
        cursor = conn.cursor()
        try:
            codigos_sem_estrutura = []
            for codigo_produto in codigos_filho:
                query_consulta_tipo_produto = f"""SELECT B1_TIPO FROM {self.database}.dbo.SB1010 WHERE 
                B1_COD = '{codigo_produto}' AND B1_TIPO IN ('PI','PA');"""
                cursor.execute(query_consulta_tipo_produto)
                resultado_tipo_produto = cursor.fetchone()
                if resultado_tipo_produto:
                    query_consulta_estrutura_totvs = f"""
                    SELECT 
                        *
                    FROM 
                        {self.database}.dbo.SG1010
                    WHERE 
                        G1_COD = '{codigo_produto}' 
                    AND 
                        G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*'
                    AND 
                        G1_REVFIM = (
                            SELECT 
                                MAX(G1_REVFIM) 
                            FROM 
                                {self.database}.dbo.SG1010 
                            WHERE 
                                G1_COD = '{codigo_produto}' 
                            AND 
                                G1_REVFIM <> 'ZZZ' 
                            AND 
                                D_E_L_E_T_ <> '*');
                    """
                    cursor.execute(query_consulta_estrutura_totvs)
                    resultado = cursor.fetchone()

                    if not resultado:
                        codigos_sem_estrutura.append(codigo_produto)

            if codigos_sem_estrutura:
                mensagem = (f"CÓDIGOS FILHO SEM ESTRUTURA NO TOTVS:\n\n{', '.join(codigos_sem_estrutura)}"
                            f"\n\nEfetue o cadastro da estrutura de cada um deles e tente novamente!\n\nツ")
                raise ValueError(mensagem)

        except pyodbc.Error as sql_ex:
            # Exibe uma caixa de diálogo se a conexão ou a consulta falhar
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(sql_ex)}",
                                             "Erro ao consultar o cadastro de estrutura dos itens filho", 16 | 0)
            raise
        except ValueError:
            raise
        except Exception as ex:
            # Trata outros erros inesperados
            ctypes.windll.user32.MessageBoxW(0, f"Erro inesperado: {str(ex)}",
                                             "Erro ao consultar o cadastro de estrutura dos itens filho", 16 | 0)
            raise
        finally:
            # Fecha a conexão com o banco de dados se estiver aberta
            if 'conn' in locals():
                conn.close()

    def remover_linhas_duplicadas_e_consolidar_quantidade(self, df_excel):
        # Agrupa o DataFrame pela combinação única de código e descrição
        grouped = df_excel.groupby([self.indice_coluna_codigo_excel, self.indice_coluna_descricao_excel])

        # Inicializa um novo DataFrame para armazenar o resultado
        df_sem_duplicatas = pd.DataFrame(columns=df_excel.columns)

        # Itera sobre os grupos consolidando as quantidades
        for _, group in grouped:
            group_filtrado = group[
                group[self.indice_coluna_dimensao].apply(lambda x: isinstance(x, (int, float)) and not pd.isna(x))]
            quantidade_consolidada = group[self.indice_coluna_quantidade_excel].sum()
            dimensao_consolidada = (group_filtrado[self.indice_coluna_quantidade_excel] * group_filtrado[
                self.indice_coluna_dimensao]).sum()
            peso_consolidado = group[self.indice_coluna_peso_excel].sum()

            # Adiciona uma linha ao DataFrame sem duplicatas
            df_sem_duplicatas = pd.concat([df_sem_duplicatas, group.head(1)])
            df_sem_duplicatas.loc[
                df_sem_duplicatas.index[-1], self.indice_coluna_quantidade_excel] = quantidade_consolidada
            df_sem_duplicatas.loc[df_sem_duplicatas.index[-1], self.indice_coluna_dimensao] = dimensao_consolidada
            df_sem_duplicatas.loc[df_sem_duplicatas.index[-1], self.indice_coluna_peso_excel] = peso_consolidado

        return df_sem_duplicatas

    def validar_codigo_filho_igual_pai(self, nome_desenho, df_excel):
        codigo_filho_diferente_codigo_pai = df_excel.iloc[:, self.indice_coluna_codigo_excel] != f"{nome_desenho}"
        if not codigo_filho_diferente_codigo_pai.all():
            ValueError("EXISTE CÓDIGO FILHO NA BOM IGUAL AO CÓDIGO PAI\n\nPor favor, corrija o código e "
                       "tente novamente!\n\nツ")

    def validacao_quantidades(self, df_excel):
        validar_quantidades = df_excel.iloc[:, self.indice_coluna_quantidade_excel].notna() & (
                df_excel.iloc[:, self.indice_coluna_quantidade_excel] != '') & (
                                      pd.to_numeric(df_excel.iloc[:, self.indice_coluna_quantidade_excel],
                                                    errors='coerce') > 0)
        if not validar_quantidades.all():
            raise ValueError("QUANTIDADE INVÁLIDA ENCONTRADA\n\nAs quantidades devem ser números, não nulas, "
                             "sem espaços em branco e maiores que zero.\nPor favor, corrija a quantidade e tente "
                             "novamente!\n\nツ")

    def validacao_pesos(self, df_excel):
        validar_pesos = df_excel.iloc[:, self.indice_coluna_peso_excel].notna() & (
                (df_excel.iloc[:, self.indice_coluna_peso_excel] == 0) |
                (pd.to_numeric(df_excel.iloc[:, self.indice_coluna_peso_excel], errors='coerce') > 0))
        if not validar_pesos.all():
            raise ValueError("PESO INVÁLIDO ENCONTRADO\n\nOs pesos devem ser números, não nulos, sem espaços "
                             "em branco e maiores ou iguais à zero.\nPor favor, corrija-os e tente novamente!\n\nツ")

    def validacao_pesos_unidade_kg(self, df_excel):
        lista_codigos_peso_zero = []
        try:

            for index, row in df_excel.iterrows():
                codigo_filho = row.iloc[self.indice_coluna_codigo_excel]
                unidade_medida = self.obter_unidade_medida_codigo_filho(codigo_filho)

                if unidade_medida == 'KG':
                    peso = row.iloc[self.indice_coluna_peso_excel]
                    if peso <= 0:
                        lista_codigos_peso_zero.append(codigo_filho)

            if lista_codigos_peso_zero:
                raise ValueError(f"PESO ZERO ENCONTRADO\n\nCertifique-se de que TODOS os pesos dos itens com "
                                 f"unidade de medida em 'kg' (kilograma) sejam MAIORES QUE ZERO."
                                 f"\n\n{lista_codigos_peso_zero}\n\nPor favor, corrija-o(s) e tente novamente!\n\nツ")
        except Exception:
            raise

    def validar_formato_campo_dimensao(self, dimensao):

        dimensao_sem_espaco = dimensao.replace(' ', '')

        if re.match(self.regex_campo_dimensao, dimensao_sem_espaco):
            return True
        else:
            return False

    def formatar_campos_dimensao(self, dataframe):
        items_mt_m2_dimensao_incorreta = {}
        items_unidade_incorreta = {}
        df_campo_dimensao_formatado = dataframe.copy()

        for i, dimensao in enumerate(df_campo_dimensao_formatado.iloc[:, self.indice_coluna_dimensao]):
            codigo_filho = df_campo_dimensao_formatado.iloc[i, self.indice_coluna_codigo_excel]
            descricao = df_campo_dimensao_formatado.iloc[i, self.indice_coluna_descricao_excel]
            unidade_de_medida = self.obter_unidade_medida_codigo_filho(codigo_filho)

            if unidade_de_medida in ('MT', 'M2', 'M3') and self.validar_formato_campo_dimensao(str(dimensao)):
                unidade_de_medida_totvs = unidade_de_medida.replace('2', '²').replace('3', '³').replace('T', '').lower()
                if unidade_de_medida_totvs != extrair_unidade_medida(str(dimensao).lower()):
                    items_unidade_incorreta[codigo_filho] = descricao
                else:
                    dimensao_final = dimensao.lower().split('m')[0].replace(',', '.')
                    if float(dimensao_final) <= 0:
                        items_mt_m2_dimensao_incorreta[codigo_filho] = descricao
                    else:
                        df_campo_dimensao_formatado.iloc[i, self.indice_coluna_dimensao] = float(dimensao_final)
            elif unidade_de_medida in ('MT', 'M2', 'M3'):
                items_mt_m2_dimensao_incorreta[codigo_filho] = descricao

        if items_mt_m2_dimensao_incorreta:
            mensagem = ''
            mensagem_fixa = f"""
            ATENÇÃO!
            VALOR DA DIMENSÃO FORA DO FORMATO PADRÃO
            
            Insira na coluna DIMENSÃO o valor
            conforme os padrões abaixo:
            
            1. Quando a unidade for METRO 'm':   
            X.XXX m ou X m
                
            2. Quando a unidade for METRO QUADRADO 'm²':     
            X.XXX m² ou X m²
            
            3. Quando a unidade for METRO CÚBICO 'm³':     
            X.XXX m³ ou X m³
                
            3. É permitido usar tanto ponto '.'
            quanto vírgula ','
                    
            4. O valor deve ser sempre maior que zero.
            
            5. (Opcional) Se precisar indicar qualquer
            informação adicional, insira entre PARÊNTESES,
            conforme o exemplo abaixo:
            
            1 m² (400x500x2)
            
            Verifique o campo DIMENSÃO do(s) código(s)
            abaixo:\n"""

            for codigo, descricao in items_mt_m2_dimensao_incorreta.items():
                mensagem += f"""
            {codigo} - {descricao[:18] + '...' if len(descricao) > 18 else descricao}"""
            raise ValueError(mensagem_fixa + mensagem)
        if items_unidade_incorreta:
            mensagem = ''
            mensagem_fixa = f"""
            OPS... Unidade de medida errada (m, m² ou m³)
            
            Provavelmente houve um erro de digitação.        
            Por favor verifique a unidade de medida do(s)
            código(s) abaixo e corrija no campo DIMENSÃO
            com a unidade correta.\n"""

            for codigo, descricao in items_unidade_incorreta.items():
                mensagem += f"""
            {codigo} - {descricao[:18] + '...' if len(descricao) > 18 else descricao}"""
            raise ValueError(mensagem_fixa + mensagem)

        return df_campo_dimensao_formatado

    def validacao_codigo_bloqueado(self, dataframe):
        conn = pyodbc.connect(
            f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};'
            f'PWD={self.password}')
        cursor = conn.cursor()
        try:
            codigos_bloqueados = {}

            for i, row in dataframe.iterrows():
                codigo = row.iloc[self.indice_coluna_codigo_excel]
                descricao = row.iloc[self.indice_coluna_descricao_excel]
                query_retorna_valor_campo_bloqueio = f"""
                SELECT 
                    B1_MSBLQL 
                FROM 
                    {self.database}.dbo.SB1010 
                WHERE 
                    B1_COD = '{codigo}' 
                AND 
                    B1_REVATU <> 'ZZZ' 
                AND 
                    D_E_L_E_T_ <> '*';
                """

                cursor.execute(query_retorna_valor_campo_bloqueio)
                resultado = cursor.fetchone()[0]

                if resultado == '1':
                    codigos_bloqueados[codigo] = descricao

            if codigos_bloqueados:
                mensagem = ''
                mensagem_fixa = f"""
        ESTRUTURA NÃO CADASTRADA
                
        Código bloqueado encontrado!
                
        Verificar:
        """
                for codigo, descricao in codigos_bloqueados.items():
                    mensagem += f"""
        {codigo} - {descricao[:18] + '...' if len(descricao) > 18 else descricao}"""
                raise ValueError(mensagem_fixa + mensagem)

        except pyodbc.Error as sql_ex:
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o banco de dados. Erro: {str(sql_ex)}",
                                             "Erro ao consultar campo BLOQUEIO (B1_MSBLQL) na tabela produtos",
                                             16 | 0)
            raise
        except ValueError:
            raise
        except Exception as ex:
            # Trata outros erros inesperados
            ctypes.windll.user32.MessageBoxW(0, f"Erro inesperado: {str(ex)}",
                                             "Erro ao consultar campo BLOQUEIO (B1_MSBLQL) na tabela produtos", 16 | 0)
            raise
        finally:
            if 'conn' in locals():
                conn.close()

    def codigo_filho_igual_ao_campo_codigo(self, dataframe):
        codigos_diferentes = {}
        codigo_nome_desenho_verificado = ''

        for index, row in dataframe.iterrows():
            codigo_filho = row.iloc[self.indice_coluna_codigo_excel].strip()
            codigo_nome_desenho = str(row.iloc[self.indice_coluna_nome_arquivo]).replace('\n', '')
            descricao = row.iloc[self.indice_coluna_descricao_excel]

            if len(codigo_nome_desenho) > 13 and codigo_nome_desenho[13] == ' ' and codigo_nome_desenho.startswith(
                    ('C', 'M', 'E')):
                codigo_nome_desenho_verificado = codigo_nome_desenho.split(' ')[0]
            elif len(codigo_nome_desenho) > 13 and codigo_nome_desenho[13] == '_' and codigo_nome_desenho.startswith(
                    ('C', 'M', 'E')):
                codigo_nome_desenho_verificado = codigo_nome_desenho.split('_')[0]
            elif len(codigo_nome_desenho) > 13 and codigo_nome_desenho[13] == '-' and codigo_nome_desenho.startswith(
                    ('C', 'M', 'E')):
                codigo_nome_desenho_verificado = codigo_nome_desenho[:13]
            elif len(codigo_nome_desenho) > 13 and codigo_nome_desenho[7] == '-':
                codigo_nome_desenho_verificado = codigo_nome_desenho.split('-')[0]
            elif len(codigo_nome_desenho) > 13 and codigo_nome_desenho[9] == '-':
                codigo_nome_desenho_verificado = codigo_nome_desenho.split('-')[0]
            elif (len(codigo_nome_desenho) == 9 or len(codigo_nome_desenho) == 10) and '.' in codigo_nome_desenho:
                codigo_nome_desenho_verificado = codigo_nome_desenho.replace('.', '')
            elif len(codigo_nome_desenho) == 8 or len(codigo_nome_desenho) == 13:
                codigo_nome_desenho_verificado = codigo_nome_desenho

            if len(codigo_nome_desenho_verificado) <= 7:
                codigo_filho_formatado = extrair_numero(codigo_filho).lstrip("0")
                if codigo_filho_formatado != codigo_nome_desenho_verificado:
                    codigos_diferentes[codigo_nome_desenho] = descricao
            elif len(codigo_nome_desenho_verificado) == 8:
                primeira_parte = codigo_nome_desenho_verificado[:3]
                segunda_parte = codigo_nome_desenho_verificado[3:5]
                terceira_parte = codigo_nome_desenho_verificado[5:9]
                codigo_nome_desenho_verificado = 'M-' + primeira_parte + '-0' + segunda_parte + '-' + terceira_parte
                if codigo_filho != codigo_nome_desenho_verificado:
                    codigos_diferentes[codigo_nome_desenho] = descricao
            elif len(codigo_nome_desenho_verificado) == 9:
                codigo_nome_desenho_verificado = 'C-00' + codigo_nome_desenho_verificado.replace('.', '-')
                if codigo_filho != codigo_nome_desenho_verificado:
                    codigos_diferentes[codigo_nome_desenho] = descricao
            elif len(codigo_nome_desenho_verificado) == 13:
                if codigo_filho != codigo_nome_desenho_verificado:
                    codigos_diferentes[codigo_nome_desenho] = descricao
            else:
                codigos_diferentes[codigo_nome_desenho] = descricao

        if codigos_diferentes:
            mensagem_codigos = ''
            mensagem_fixa = (f"\n\nATENÇÃO\n\n O CÓDIGO FILHO PODE ESTAR ERRADO.\n\n"
                             f"---> APÓS A CORREÇÃO ATUALIZE O TEMPLATE DA BOM <---\n\n"
                             f"O nome do desenho deve ser igual ao campo Nº DA PEÇA do formulário\n\nVerificar:\n\n")
            mensagem_final = "\nEUREKA®"
            i = 1
            for code, description in codigos_diferentes.items():
                if code == 'nan':
                    mensagem_codigos += (f"{i}. Alterar CONFIG. NOME USUÁRIO nas propriedades da peça/montagem"
                                         f"- {description[:20] + '...' if len(description) > 20 else description}\n\n")
                else:
                    mensagem_codigos += (f"{i}. {code[:18] + '...' if len(code) > 18 else code} "
                                         f"- {description[:20] + '...' if len(description) > 20 else description}\n\n")
                i += 1
            raise Exception(mensagem_fixa + mensagem_codigos + mensagem_final)

    def validacao_de_dados_bom(self, excel_file_path):
        try:
            df_excel = pd.read_excel(excel_file_path, sheet_name='Planilha1', header=None)
            pos_index = df_excel[df_excel.iloc[:, 0] == "Pos."].index[0]
            df_excel = df_excel.iloc[:pos_index]

            df_excel.loc[:, self.indice_coluna_codigo_excel] = df_excel.loc[:, self.indice_coluna_codigo_excel].apply(
                lambda x: x.strip())

            tipo_da_bom = verificar_se_template_bom_esta_correto(df_excel)
            self.validar_formato_codigos_filho(df_excel)
            self.validacao_quantidades(df_excel)
            validar_descricao(df_excel.iloc[:, self.indice_coluna_descricao_excel])
            self.validacao_pesos(df_excel)
            self.validar_codigo_filho_igual_pai(self.nome_desenho, df_excel)

            if tipo_da_bom == "montagem":
                self.codigo_filho_igual_ao_campo_codigo(df_excel)

            self.verificar_cadastro_codigo_filho(df_excel.iloc[:, self.indice_coluna_codigo_excel].tolist())

            df_excel_campo_dimensao_tratado = self.formatar_campos_dimensao(df_excel)

            df_bom_excel = self.remover_linhas_duplicadas_e_consolidar_quantidade(
                df_excel_campo_dimensao_tratado)
            df_bom_excel.iloc[:, self.indice_coluna_codigo_excel] = (
                df_bom_excel.iloc[:, self.indice_coluna_codigo_excel].str.strip())

            self.verificar_codigo_repetido(df_bom_excel)
            self.validacao_codigo_bloqueado(df_excel)
            self.verificar_se_existe_estrutura_codigos_filho(
                df_bom_excel.iloc[:, self.indice_coluna_codigo_excel].tolist())
            self.validacao_pesos_unidade_kg(df_excel)

            return df_bom_excel

        except Exception as e:
            raise Exception(f"Erro durante a validação da BOM\n\n{str(e)}")

    def atualizar_campo_revisao_do_codigo_pai(self, codigo_pai, numero_revisao):
        query_atualizar_campo_revisao = f"""
        UPDATE 
            {self.database}.dbo.SB1010 
        SET 
            B1_REVATU = N'{numero_revisao}' 
        WHERE 
            B1_COD = N'{codigo_pai}' 
        AND 
            D_E_L_E_T_ <> '*';"""
        try:
            # Uso do Context Manager para garantir o fechamento adequado da conexão
            with pyodbc.connect(
                    f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};'
                    f'PWD={self.password}') as conn:
                cursor = conn.cursor()
                cursor.execute(query_atualizar_campo_revisao)

        except pyodbc.Error as sql_ex:
            # Trata erros relacionados ao SQL
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(sql_ex)}",
                                             "Erro ao atualizar o campo REV. do código pai", 16 | 0)
            raise
        except Exception as ex:
            # Trata outros erros inesperados
            ctypes.windll.user32.MessageBoxW(0, f"Erro inesperado: {str(ex)}",
                                             "Erro ao atualizar o campo REV. do código pai", 16 | 0)
            raise

    def atualizar_campo_data_ultima_revisao_do_codigo_pai(self, codigo_pai):
        data_atual = formatar_data_atual()
        query_atualizar_data_ultima_revisao = f"""
            UPDATE 
                {self.database}.dbo.SB1010 
            SET 
                B1_UREV = N'{data_atual}' 
            WHERE 
                B1_COD = N'{codigo_pai}' 
            AND 
                D_E_L_E_T_ <> '*';"""
        try:
            with pyodbc.connect(
                    f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};'
                    f'PWD={self.password}') as conn:
                cursor = conn.cursor()
                cursor.execute(query_atualizar_data_ultima_revisao)
                return True
        except pyodbc.Error as sql_ex:
            # Trata erros relacionados ao SQL
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(sql_ex)}",
                                             "Erro ao atualizar o campo DATA ULT. REV. do código pai", 16 | 0)
            raise
        except Exception as ex:
            # Trata outros erros inesperados
            ctypes.windll.user32.MessageBoxW(0, f"Erro inesperado: {str(ex)}",
                                             "Erro ao atualizar o campo DATA ULT. REV. do código pai", 16 | 0)
            raise

    def verificar_estrutura_codigo_pai(self, codigo_pai):
        # Tente estabelecer a conexão com o banco de dados
        conn_str = (f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};'
                    f'PWD={self.password}')
        engine = create_engine(f'mssql+pyodbc:///?odbc_connect={conn_str}')
        try:
            query = f"""
                SELECT 
                    * 
                FROM 
                    {self.database}.dbo.SG1010
                WHERE 
                    G1_COD = '{codigo_pai}' 
                AND 
                    G1_REVFIM <> 'ZZZ'
                AND 
                    D_E_L_E_T_ <> '*'
                AND 
                    G1_REVFIM = (
                        SELECT 
                            MAX(G1_REVFIM)
                        FROM 
                            {self.database}.dbo.SG1010
                        WHERE 
                            G1_COD = '{codigo_pai}' 
                        AND 
                            G1_REVFIM <> 'ZZZ' 
                        AND 
                            D_E_L_E_T_ <> '*'
                    );
            """

            # Executa a query SELECT e obtém os resultados em um DataFrame
            result = pd.DataFrame()
            self.pai_tem_estrutura = pd.read_sql(query, engine)

        except SQLAlchemyError as sql_ex:
            # Trata erros relacionados ao SQL
            ctypes.windll.user32.MessageBoxW(0, f"Erro de banco de dados: {str(sql_ex)}",
                                             "Erro ao verificar estrutura no TOTVS", 16 | 0)
            raise
        except Exception as ex:
            # Trata outros erros inesperados
            ctypes.windll.user32.MessageBoxW(0, f"Erro inesperado: {str(ex)}",
                                             "Erro ao verificar estrutura no TOTVS", 16 | 0)
            raise
        finally:
            # Fecha a conexão com o banco de dados se estiver aberta
            if 'engine' in locals():
                engine.dispose()

    def obter_ultima_pk_tabela_estrutura(self):
        query_ultima_pk_tabela_estrutura = f"""
            SELECT 
                TOP 1 R_E_C_N_O_ 
            FROM 
                {self.database}.dbo.SG1010 
            ORDER BY 
                R_E_C_N_O_ DESC;"""
        try:
            # Uso do Context Manager para garantir o fechamento adequado da conexão
            with pyodbc.connect(
                    f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};'
                    f'PWD={self.password}') as conn:
                cursor = conn.cursor()
                cursor.execute(query_ultima_pk_tabela_estrutura)
                resultado_ultima_pk_tabela_estrutura = cursor.fetchone()
                valor_ultima_pk = resultado_ultima_pk_tabela_estrutura[0]

                return valor_ultima_pk

        except pyodbc.Error as sql_ex:
            # Tratar erros relacionados ao SQL
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(sql_ex)}",
                                             "Erro ao obter última PK da tabela estrutura", 16 | 0)
            raise
        except Exception as ex:
            # Trata outros erros inesperados
            ctypes.windll.user32.MessageBoxW(0, f"Erro inesperado: {str(ex)}",
                                             "Erro ao obter última PK da tabela estrutura", 16 | 0)
            raise

    def obter_revisao_codigo_pai(self, codigo_pai, primeiro_cadastro):
        query_revisao_inicial = f"""
        SELECT 
            B1_REVATU 
        FROM 
            {self.database}.dbo.SB1010 
        WHERE 
            B1_COD = '{codigo_pai}'"""
        try:
            # Uso do Context Manager para garantir o fechamento adequado da conexão
            with pyodbc.connect(
                    f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};'
                    f'PWD={self.password}') as conn:
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

        except pyodbc.Error as sql_ex:
            # Tratar erros relacionados ao SQL
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(sql_ex)}",
                                             "Erro ao consultar revisão do código pai na tabela de produtos", 16 | 0)
            raise
        except Exception as ex:
            # Trata outros erros inesperados
            ctypes.windll.user32.MessageBoxW(0, f"Erro inesperado: {str(ex)}",
                                             "Erro ao consultar revisão do código pai na tabela de produtos", 16 | 0)
            raise

    def obter_unidade_medida_codigo_filho(self, codigo_filho):
        query_unidade_medida_codigo_filho = f"""
            SELECT 
                B1_UM 
            FROM 
                {self.database}.dbo.SB1010 
            WHERE 
                B1_COD = '{codigo_filho}'"""
        try:
            with pyodbc.connect(
                    f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};'
                    f'PWD={self.password}') as conn:
                cursor = conn.cursor()
                cursor.execute(query_unidade_medida_codigo_filho)

                unidade_medida = cursor.fetchone()
                valor_unidade_medida = unidade_medida[0]

                return valor_unidade_medida

        except pyodbc.Error as sql_ex:
            # Trata erros relacionados ao SQL
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(sql_ex)}",
                                             "Erro ao consultar a unidade de medida de código filho", 16 | 0)
            raise
        except Exception as ex:
            # Trata outros erros inesperados
            ctypes.windll.user32.MessageBoxW(0, f"Erro inesperado: {str(ex)}",
                                             "Erro ao consultar a unidade de medida de código filho", 16 | 0)
            raise

    def verificar_cadastro_codigo_pai(self, codigo_pai):
        query_consulta_produto_codigo_pai = f"""
            SELECT 
                B1_COD 
            FROM 
                {self.database}.dbo.SB1010 
            WHERE 
                B1_COD = '{codigo_pai}'"""
        try:
            with pyodbc.connect(
                    f'DRIVER='
                    f'{self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};'
                    f'PWD={self.password}') as conn:
                cursor = conn.cursor()
                cursor.execute(query_consulta_produto_codigo_pai)
                resultado = cursor.fetchone()

                if resultado:
                    return True
                else:
                    message = f"O cadastro do item pai não foi encontrado!"
                    raise ValueError(message)

        except pyodbc.Error as sql_ex:
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(sql_ex)}",
                                             "Erro ao consultar cadastro do código pai", 16 | 0)
            raise
        except Exception:
            raise

    def criar_nova_estrutura_totvs(self, codigo_pai, df_bom_excel, revisao_inicial):
        codigo_filho, quantidade, unidade_medida, ultima_pk_tabela_estrutura = '', '', '', ''
        conn = pyodbc.connect(
            f'DRIVER='
            f'{self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}')
        cursor = conn.cursor()
        try:
            ultima_pk_tabela_estrutura = self.obter_ultima_pk_tabela_estrutura()
            data_atual_formatada = formatar_data_atual()
            for index, row in df_bom_excel.iterrows():
                ultima_pk_tabela_estrutura += 1
                codigo_filho = row.iloc[self.indice_coluna_codigo_excel]
                quantidade = row.iloc[self.indice_coluna_quantidade_excel]
                unidade_medida = self.obter_unidade_medida_codigo_filho(codigo_filho)

                if unidade_medida == 'KG':
                    quantidade = row.iloc[self.indice_coluna_peso_excel]
                elif unidade_medida in ('MT', 'M2', 'M3'):
                    quantidade = row.iloc[self.indice_coluna_dimensao]

                quantidade_formatada = "{:.2f}".format(float(quantidade))

                query_criar_nova_estrutura_totvs = f"""
                    INSERT INTO {self.database}.dbo.SG1010 
                    (G1_FILIAL, G1_COD, G1_COMP, G1_TRT, G1_XUM, G1_QUANT, G1_PERDA, G1_INI, G1_FIM, G1_OBSERV, 
                    G1_FIXVAR, G1_GROPC, G1_OPC, G1_REVINI, G1_NIV, G1_NIVINV, G1_REVFIM, 
                    G1_OK, G1_POTENCI, G1_TIPVEC, G1_VECTOR, G1_VLCOMPE, G1_LOCCONS, G1_USAALT, G1_FANTASM, G1_LISTA, 
                    D_E_L_E_T_, R_E_C_N_O_, R_E_C_D_E_L_) 
                    VALUES (N'0101', N'{codigo_pai}  ', N'{codigo_filho}  ', N'   ', N'{unidade_medida}', 
                    {quantidade_formatada}, 0.0, N'{data_atual_formatada}', N'20491231', 
                    N'                                             ', N'V', N'   ', N'    ', N'{revisao_inicial}', 
                    N'01', N'99', N'{revisao_inicial}', N'    ', 0.0, N'      ', N'      ', 
                    N'N', N'  ', N'1', N' ', N'          ', 
                    N' ', {ultima_pk_tabela_estrutura}, 0);
                """

                cursor.execute(query_criar_nova_estrutura_totvs)
            conn.commit()

        except pyodbc.Error as sql_ex:
            raise Exception(f"Falha ao criar nova estrutura. Erro de banco de dados: {str(sql_ex)} - "
                            f"PK-{ultima_pk_tabela_estrutura} - {codigo_pai} - {codigo_filho} - "
                            f"{quantidade} - {unidade_medida}")
        except Exception:
            raise
        finally:
            cursor.close()
            conn.close()

    def atualizar_revisao_quantidade_totvs(self, codigo_pai, itens_em_comum, revisao_atualizada, revisao_anterior):
        conn = pyodbc.connect(
            f'DRIVER='
            f'{self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}')
        cursor = conn.cursor()
        try:
            for index, row in itens_em_comum.iterrows():
                codigo_filho = row.iloc[self.indice_coluna_codigo_excel]
                quantidade = row.iloc[self.indice_coluna_quantidade_excel]
                unidade_medida = self.obter_unidade_medida_codigo_filho(codigo_filho)

                if unidade_medida == 'KG':
                    quantidade = row.iloc[self.indice_coluna_peso_excel]
                elif unidade_medida in ('MT', 'M2', 'M3'):
                    quantidade = row.iloc[self.indice_coluna_dimensao]
                quantidade_formatada = "{:.2f}".format(float(quantidade))

                query = f"""
                    UPDATE 
                        {self.database}.dbo.SG1010 
                    SET 
                        G1_QUANT = {quantidade_formatada},
                        G1_REVFIM = '{revisao_atualizada}'
                    WHERE 
                        G1_COD = '{codigo_pai}'
                    AND 
                        G1_COMP = '{codigo_filho}'
                    AND 
                        G1_REVFIM <> 'ZZZ'
                    AND 
                        D_E_L_E_T_ <> '*'
                    AND 
                        G1_REVFIM = '{revisao_anterior}';
                """

                cursor.execute(query)
            conn.commit()

        except pyodbc.Error as sql_ex:
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou consulta. Erro: {str(sql_ex)}",
                                             "Erro ao atualizar itens já existentes da estrutura", 16 | 0)
            raise
        except Exception:
            raise
        finally:
            cursor.close()
            conn.close()

    def atualizar_quantidade_totvs(self, codigo_pai, itens_em_comum, revisao_anterior):
        conn = pyodbc.connect(
            f'DRIVER='
            f'{self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}')
        cursor = conn.cursor()
        try:
            for index, row in itens_em_comum.iterrows():
                codigo_filho = row.iloc[self.indice_coluna_codigo_excel]
                quantidade = row.iloc[self.indice_coluna_quantidade_excel]
                unidade_medida = self.obter_unidade_medida_codigo_filho(codigo_filho)

                if unidade_medida == 'KG':
                    quantidade = row.iloc[self.indice_coluna_peso_excel]
                elif unidade_medida in ('MT', 'M2', 'M3'):
                    quantidade = row.iloc[self.indice_coluna_dimensao]
                quantidade_formatada = "{:.2f}".format(float(quantidade))

                query = f"""
                    UPDATE 
                        {self.database}.dbo.SG1010
                    SET 
                        G1_QUANT = {quantidade_formatada}
                    WHERE 
                        G1_COD = '{codigo_pai}'
                    AND 
                        G1_COMP = '{codigo_filho}'
                    AND 
                        G1_REVFIM <> 'ZZZ'
                    AND 
                        D_E_L_E_T_ <> '*'
                    AND 
                        G1_REVFIM = '{revisao_anterior}';
                """

                cursor.execute(query)
            conn.commit()

        except pyodbc.Error as sql_ex:
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou consulta. Erro: {str(sql_ex)}",
                                             "Erro ao atualizar itens já existentes da estrutura", 16 | 0)
            raise
        except Exception:
            raise
        finally:
            cursor.close()
            conn.close()

    def inserir_itens_estrutura_totvs(self, codigo_pai, df_itens_adicionados,
                                      revisao_atualizada_estrutura):
        conn = pyodbc.connect(f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};'
                              f'PWD={self.password}')
        cursor = conn.cursor()
        try:
            ultima_pk_tabela_estrutura = self.obter_ultima_pk_tabela_estrutura()
            data_atual_formatada = formatar_data_atual()

            revisao_inicial = revisao_atualizada_estrutura

            for index, row in df_itens_adicionados.iterrows():
                ultima_pk_tabela_estrutura += 1
                codigo_filho = row.iloc[self.indice_coluna_codigo_excel]
                quantidade = row.iloc[self.indice_coluna_quantidade_excel]
                unidade_medida = self.obter_unidade_medida_codigo_filho(codigo_filho)

                if unidade_medida == 'KG':
                    quantidade = row.iloc[self.indice_coluna_peso_excel]
                elif unidade_medida in ('MT', 'M2', 'M3'):
                    quantidade = row.iloc[self.indice_coluna_dimensao]
                quantidade_formatada = "{:.2f}".format(float(quantidade))

                query_criar_nova_estrutura_totvs = f"""
                    INSERT INTO {self.database}.dbo.SG1010 
                    (G1_FILIAL, G1_COD, G1_COMP, G1_TRT, G1_XUM, G1_QUANT, G1_PERDA, G1_INI, G1_FIM, G1_OBSERV, 
                    G1_FIXVAR, G1_GROPC, G1_OPC, G1_REVINI, G1_NIV, G1_NIVINV, G1_REVFIM, 
                    G1_OK, G1_POTENCI, G1_TIPVEC, G1_VECTOR, G1_VLCOMPE, G1_LOCCONS, G1_USAALT, G1_FANTASM, 
                    G1_LISTA, D_E_L_E_T_, R_E_C_N_O_, R_E_C_D_E_L_) 
                    VALUES (N'0101', N'{codigo_pai}  ', N'{codigo_filho}  ', N'   ', N'{unidade_medida}', 
                    {quantidade_formatada}, 0.0, N'{data_atual_formatada}', N'20491231', 
                    N'                                             ', N'V', N'   ', N'    ', N'{revisao_inicial}', 
                    N'01', N'99', N'{revisao_atualizada_estrutura}', N'    ', 0.0, N'      ', N'      ', 
                    N'N', N'  ', N'1', N' ', N'          ', 
                    N' ', {ultima_pk_tabela_estrutura}, 0);
                """

                cursor.execute(query_criar_nova_estrutura_totvs)
            conn.commit()

        except pyodbc.Error as sql_ex:
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou consulta. Erro: {str(sql_ex)}",
                                             "Erro ao inserir novos itens na estrutura", 16 | 0)
            raise
        except Exception:
            raise
        finally:
            cursor.close()
            conn.close()

    def comparar_bom_com_totvs(self, df_bom_excel, resultado_query_consulta_estrutura_totvs):
        resultado_query_consulta_estrutura_totvs['G1_COMP'] = resultado_query_consulta_estrutura_totvs[
            'G1_COMP'].str.strip()

        # Códigos em comum
        self.itens_em_comum = df_bom_excel[
            df_bom_excel.iloc[:, self.indice_coluna_codigo_excel].isin(
                resultado_query_consulta_estrutura_totvs['G1_COMP'])].copy()

        # Códigos adicionados no BOM
        self.itens_adicionados = df_bom_excel[
            ~df_bom_excel.iloc[:, self.indice_coluna_codigo_excel].isin(
                resultado_query_consulta_estrutura_totvs['G1_COMP'])].copy()

        # Códigos removidos no BOM
        self.itens_removidos = resultado_query_consulta_estrutura_totvs[
            ~resultado_query_consulta_estrutura_totvs['G1_COMP'].isin(
                df_bom_excel.iloc[:, self.indice_coluna_codigo_excel])].copy()

        return self.itens_em_comum, self.itens_adicionados, self.itens_removidos

    def start_task(self):
        thread = threading.Thread(target=self.executar_logica)
        thread.start()

    def update_progress(self, value):
        self.progress['value'] = value
        root.update_idletasks()

    def executar_logica(self):
        status_processo = None
        mensagem_processo = {
            "cancelado": '❌ Processo cancelado!',
            "sucesso": '✔️ Processo finalizado com sucesso!'
        }
        excel_file_path = obter_caminho_arquivo_excel(self.nome_desenho)
        try:
            delay = 0.4
            self.status_label.config(text="Iniciando cadastro...")
            self.update_progress(5)
            time.sleep(0.7)

            self.status_label.config(text="Validando formato do código pai...")
            self.update_progress(8)
            time.sleep(delay)
            formato_codigo_pai_correto = self.validar_formato_codigo_pai(self.nome_desenho)

            existe_cadastro_codigo_pai = False
            if formato_codigo_pai_correto:
                self.status_label.config(text="Verificando cadastro do código pai...")
                self.update_progress(10)
                time.sleep(delay)
                existe_cadastro_codigo_pai = self.verificar_cadastro_codigo_pai(self.nome_desenho)

            if formato_codigo_pai_correto and existe_cadastro_codigo_pai:
                self.status_label.config(text="Validando dados da tabela de BOM...")
                self.update_progress(20)
                time.sleep(delay)
                df_bom_excel = self.validacao_de_dados_bom(excel_file_path)

                # VEFIFICA SE BOM ESTÁ VAZIA
                if df_bom_excel.empty:
                    exibir_mensagem(self.titulo_janela,
                                    f"OPS!\n\nA BOM está vazia!\n\nPor gentileza, preencha adequadamente a "
                                    f"BOM e tente novamente!\n\n{self.nome_desenho}\n\nツ\n\nEUREKA®",
                                    "warning")
                    status_processo = mensagem_processo['cancelado']
                    self.update_progress(80)
                    raise Exception("BOM vazia")

                self.status_label.config(text="Verificando se já existe estrutura cadastrada...")
                self.update_progress(25)
                time.sleep(delay)
                self.verificar_estrutura_codigo_pai(self.nome_desenho)

                # CADASTRO DE NOVA ESTRUTURA NO TOTVS
                if not df_bom_excel.empty and self.pai_tem_estrutura.empty:
                    primeiro_cadastro = True
                    revisao_inicial = self.obter_revisao_codigo_pai(self.nome_desenho, primeiro_cadastro)
                    self.status_label.config(text="Cadastrando nova estrutura...")
                    self.update_progress(50)
                    time.sleep(delay)
                    self.criar_nova_estrutura_totvs(self.nome_desenho, df_bom_excel, revisao_inicial)
                    self.status_label.config(text="Atualizando revisão do código pai...")
                    self.update_progress(80)
                    time.sleep(delay)
                    self.atualizar_campo_revisao_do_codigo_pai(self.nome_desenho, revisao_inicial)
                    self.atualizar_campo_data_ultima_revisao_do_codigo_pai(self.nome_desenho)
                    status_processo = mensagem_processo['sucesso']
                    exibir_mensagem(self.titulo_janela,
                                    f"Estrutura cadastrada com sucesso!\n\n{self.nome_desenho}\n\n( ͡° ͜ʖ ͡°)\n\nEUREKA®",
                                    "info")

                elif not df_bom_excel.empty and not self.pai_tem_estrutura.empty:
                    mensagem = (f"ESTRUTURA EXISTENTE\n\nJá existe uma estrutura cadastrada no TOTVS para este produto!"
                                f"\n\n{self.nome_desenho}\n\nDeseja realizar a atualização da estrutura?")
                    usuario_quer_alterar = exibir_janela_mensagem_opcao(self.titulo_janela, mensagem)
                    self.update_progress(50)

                    if usuario_quer_alterar:
                        resultado = self.comparar_bom_com_totvs(df_bom_excel, self.pai_tem_estrutura)
                        itens_em_comum, itens_adicionados, itens_removidos = resultado
                        self.status_label.config(text="Analisando se houve mudanças na estrutura...")
                        self.update_progress(60)
                        time.sleep(delay)
                        primeiro_cadastro = False
                        revisao_atualizada = self.obter_revisao_codigo_pai(self.nome_desenho, primeiro_cadastro)
                        revisao_anterior = calculo_revisao_anterior(revisao_atualizada)

                        # ADICIONA OS NOVOS ITENS DA BOM NA ESTRUTURA NO TOTVS
                        if not itens_adicionados.empty:
                            self.status_label.config(text="Adicionando novos itens na estrutura...")
                            self.update_progress(70)
                            time.sleep(delay)
                            self.inserir_itens_estrutura_totvs(
                                self.nome_desenho, itens_adicionados, revisao_atualizada)

                        # VERIFICA SE FOI ADICIONADO OU REMOVIDO ITENS NA BOM.
                        # SE SIM, ATUALIZA A REVISÃO E AS QUANTIDADES
                        # SE NÃO, ATUALIZA APENAS AS QUANTIDADES
                        if not itens_adicionados.empty or not itens_removidos.empty:
                            if not itens_em_comum.empty:
                                self.status_label.config(text="Atualizando a revisão e as quantidades...")
                                self.update_progress(75)
                                time.sleep(delay)
                                self.atualizar_revisao_quantidade_totvs(self.nome_desenho, itens_em_comum, revisao_atualizada, revisao_anterior)

                            self.status_label.config(text="Atualizando revisão do código pai...")
                            self.update_progress(80)
                            time.sleep(delay)

                            self.atualizar_campo_revisao_do_codigo_pai(self.nome_desenho, revisao_atualizada)
                            self.atualizar_campo_data_ultima_revisao_do_codigo_pai(self.nome_desenho)

                            self.status_label.config(text="Atualização de estrutura finalizada!")
                            self.update_progress(90)
                            time.sleep(delay)
                        else:
                            self.atualizar_quantidade_totvs(self.nome_desenho, itens_em_comum, revisao_anterior)
                            self.status_label.config(text="Atualização das quantidades finalizada!")
                            self.update_progress(90)
                            time.sleep(delay)

                        exibir_mensagem(self.titulo_janela,
                                        f"Atualização de estrutura realizada com sucesso!"
                                        f"\n\n{self.nome_desenho}\n\n( ͡° ͜ʖ ͡°)\n\nEUREKA®",
                                        "info")
                        status_processo = mensagem_processo['sucesso']
                    else:
                        status_processo = mensagem_processo['cancelado']

        except Exception as e:
            exibir_mensagem(self.titulo_janela, f'Falha ao cadastrar BOM\n\n{e}', 'warning')
            status_processo = mensagem_processo['cancelado']
        finally:
            excluir_arquivo_excel_bom(excel_file_path)
            end_time = time.time()
            elapsed = end_time - self.start_time
            self.status_label.config(text=f"{status_processo}\n\n{elapsed:.3f} segundos\n\nEUREKA®")
            self.update_progress(100)


if __name__ == "__main__":
    root = tk.Tk()
    cadastro = CadastrarBomTOTVS(root)
    cadastro.start_task()
    root.attributes('-topmost', True)
    root.geometry("400x200")
    root.mainloop()
