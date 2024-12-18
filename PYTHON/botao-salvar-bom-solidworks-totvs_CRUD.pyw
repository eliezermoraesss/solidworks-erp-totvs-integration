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


class CadastrarBomTOTVS:
    def __init__(self, root):
        # Leitura dos parâmetros de conexão com o banco de dados SQL Server
        self.username, self.password, self.database, self.server = self.setup_mssql()
        self.driver = '{SQL Server}'

        root.title("Monitor de progresso")
        self.start_time = time.time()

        self.progress = ttk.Progressbar(root, orient="horizontal", length="300", mode="determinate")
        self.progress.pack(pady=20)

        self.status_label = tk.Label(root, text="")
        self.status_label.pack(pady=30)

        self.titulo_janela = "CADASTRO DE ESTRUTURA TOTVS®"

        # Arrays para armazenar os códigos
        self.codigos_adicionados_bom = []  # ITENS ADICIONADOS
        self.codigos_removidos_bom = []  # ITENS REMOVIDOS
        self.codigos_em_comum = []  # ITENS EM COMUM

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

        self.nome_desenho = self.ler_variavel_ambiente_codigo_desenho()

    def setup_mssql(self):
        caminho_do_arquivo = r"\\192.175.175.4\f\INTEGRANTES\ELIEZER\PROJETO SOLIDWORKS TOTVS\libs-python\user-password-mssql\USER_PASSWORD_MSSQL_PROD.txt"
        try:
            with open(caminho_do_arquivo, 'r') as arquivo:
                string_lida = arquivo.read()
                username, password, database, server = string_lida.split(';')
                return username, password, database, server

        except FileNotFoundError:
            ctypes.windll.user32.MessageBoxW(0,
                                             f"Erro ao ler credenciais de acesso ao banco de dados MSSQL.\n\nBase de dados ERP TOTVS PROTHEUS.\n\nPor favor, informe ao desenvolvedor/TI sobre o erro exibido.\n\nTenha um bom dia! ツ",
                                             "CADASTRO DE ESTRUTURA - TOTVS®", 16 | 0)
            sys.exit()

        except Exception as e:
            ctypes.windll.user32.MessageBoxW(0, f"Ocorreu um erro ao ler o arquivo:", "CADASTRO DE ESTRUTURA - TOTVS®",
                                             16 | 0)
            sys.exit()

    def validar_formato_codigo_pai(self, codigo_pai):
        codigo_pai_validado = any(re.match(formato, str(codigo_pai)) for formato in self.formatos_codigo)

        if not codigo_pai_validado:
            self.exibir_mensagem(self.titulo_janela,
                                 f"Este desenho foi salvo com o código fora do formato padrão ENAPLIC.\n\n{codigo_pai}\n\nPor favor renomear e salvar o desenho com o código no formato padrão.\n\nツ",
                                 "info")

        return codigo_pai_validado

    def validar_formato_codigos_filho(self, df_excel):
        validacao_codigos = df_excel.iloc[:, self.indice_coluna_codigo_excel].str.strip().apply(
            lambda x: any(re.match(formato, str(x)) for formato in self.formatos_codigo))
        if not validacao_codigos.all():
            self.exibir_mensagem(self.titulo_janela,
                                 "CÓDIGO FILHO FORA DO FORMATO PADRÃO ENAPLIC\n\nPor favor, corrija o código e tente novamente!\n\nツ",
                                 "info")
        return validacao_codigos

    def ler_variavel_ambiente_codigo_desenho(self):
        return os.getenv('CODIGO_DESENHO')

    def obter_caminho_arquivo_excel(self, codigo_desenho):
        base_path = os.environ.get('TEMP')
        return os.path.join(base_path, codigo_desenho + '.xlsx')

    def excluir_arquivo_excel_bom(self, excel_file_path):
        if os.path.exists(excel_file_path):
            os.remove(excel_file_path)

    def verificar_codigo_repetido(self, df_excel):

        codigos_repetidos = df_excel.iloc[:, self.indice_coluna_codigo_excel][
            df_excel.iloc[:, self.indice_coluna_codigo_excel].duplicated()]

        # Exibe uma mensagem se houver códigos repetidos
        if not codigos_repetidos.empty:
            self.exibir_mensagem(self.titulo_janela,
                                 f"PRODUTOS REPETIDOS NA BOM.\n\nOs códigos são iguais com descrições diferentes:\n\n{codigos_repetidos.tolist()}\n\nCorrija-os ou exclue da tabela e tente novamente!\n\nツ",
                                 "warning")
            return True
        else:
            return False

    def verificar_cadastro_codigo_filho(self, codigos_filho):
        try:
            conn = pyodbc.connect(
                f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}')
            cursor = conn.cursor()

            codigos_sem_cadastro = []

            for codigo_produto in codigos_filho:
                query_consulta_produto = f"""
                SELECT B1_COD FROM {self.database}.dbo.SB1010 WHERE B1_COD = '{codigo_produto.strip()}' AND D_E_L_E_T_ <> '*';
                """

                cursor.execute(query_consulta_produto)
                resultado = cursor.fetchone()

                if not resultado:
                    codigos_sem_cadastro.append(codigo_produto)

            if codigos_sem_cadastro:
                mensagem = f"CÓDIGOS FILHO SEM CADASTRO NO TOTVS:\n\n{', '.join(codigos_sem_cadastro)}\n\nEfetue o cadastro e tente novamente!\n\nツ"
                self.exibir_mensagem(self.titulo_janela, mensagem, "info")
                return False
            else:
                return True

        except Exception as ex:
            # Exibe uma caixa de diálogo se a conexão ou a consulta falhar
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}",
                                             "Erro ao consultar o cadastro de produtos", 16 | 0)

        finally:
            # Fecha a conexão com o banco de dados se estiver aberta
            if 'conn' in locals():
                conn.close()

    def verificar_se_existe_estrutura_codigos_filho(self, codigos_filho):
        try:
            conn = pyodbc.connect(
                f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}')
            cursor = conn.cursor()

            codigos_sem_estrutura = []

            for codigo_produto in codigos_filho:

                query_consulta_tipo_produto = f"""SELECT B1_TIPO FROM {self.database}.dbo.SB1010 WHERE B1_COD = '{codigo_produto}' AND B1_TIPO IN ('PI','PA');"""

                cursor.execute(query_consulta_tipo_produto)
                resultado_tipo_produto = cursor.fetchone()

                if resultado_tipo_produto:

                    query_consulta_estrutura_totvs = f"""SELECT *
                        FROM {self.database}.dbo.SG1010
                        WHERE G1_COD = '{codigo_produto}' 
                        AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*'
                        AND G1_REVFIM = (SELECT MAX(G1_REVFIM) FROM {self.database}.dbo.SG1010 WHERE G1_COD = '{codigo_produto}'AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*');
                    """

                    cursor.execute(query_consulta_estrutura_totvs)
                    resultado = cursor.fetchone()

                    if not resultado:
                        codigos_sem_estrutura.append(codigo_produto)

            if codigos_sem_estrutura:
                mensagem = f"CÓDIGOS FILHO SEM ESTRUTURA NO TOTVS:\n\n{', '.join(codigos_sem_estrutura)}\n\nEfetue o cadastro da estrutura de cada um deles e tente novamente!\n\nツ"
                self.exibir_mensagem(self.titulo_janela, mensagem, "info")
                return False
            else:
                return True

        except Exception as ex:
            # Exibe uma caixa de diálogo se a conexão ou a consulta falhar
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}",
                                             "Erro ao consultar o cadastro de estrutura dos itens filho", 16 | 0)

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

    def validar_descricao(self, descricoes):
        validar_descricoes = descricoes.notna() & (descricoes != '') & (descricoes.astype(str).str.strip() != '')
        if not validar_descricoes.all():
            self.exibir_mensagem(self.titulo_janela,
                                 "DESCRIÇÃO INVÁLIDA ENCONTRADA\n\nAs descrições não podem ser nulas, vazias ou conter apenas espaços em branco.\nPor favor, corrija a descrição e tente novamente!\n\nツ",
                                 "info")
        return validar_descricoes

    def verificar_codigo_filho_diferente_codigo_pai(self, nome_desenho, df_excel):
        codigo_filho_diferente_codigo_pai = df_excel.iloc[:, self.indice_coluna_codigo_excel] != f"{nome_desenho}"
        if not codigo_filho_diferente_codigo_pai.all():
            self.exibir_mensagem(self.titulo_janela,
                                 "EXISTE CÓDIGO FILHO NA BOM IGUAL AO CÓDIGO PAI\n\nPor favor, corrija o código e tente novamente!\n\nツ",
                                 "info")
        return codigo_filho_diferente_codigo_pai

    def validacao_quantidades(self, df_excel):
        validar_quantidades = df_excel.iloc[:, self.indice_coluna_quantidade_excel].notna() & (
                    df_excel.iloc[:, self.indice_coluna_quantidade_excel] != '') & (
                                          pd.to_numeric(df_excel.iloc[:, self.indice_coluna_quantidade_excel],
                                                        errors='coerce') > 0)
        if not validar_quantidades.all():
            self.exibir_mensagem(self.titulo_janela,
                                 "QUANTIDADE INVÁLIDA ENCONTRADA\n\nAs quantidades devem ser números, não nulas, sem espaços em branco e maiores que zero.\nPor favor, corrija a quantidade e tente novamente!\n\nツ",
                                 "info")
        return validar_quantidades

    def validacao_pesos(self, df_excel):
        validar_pesos = df_excel.iloc[:, self.indice_coluna_peso_excel].notna() & (
                    (df_excel.iloc[:, self.indice_coluna_peso_excel] == 0) | (
                        pd.to_numeric(df_excel.iloc[:, self.indice_coluna_peso_excel], errors='coerce') > 0))
        if not validar_pesos.all():
            self.exibir_mensagem(self.titulo_janela,
                                 "PESO INVÁLIDO ENCONTRADO\n\nOs pesos devem ser números, não nulos, sem espaços em branco e maiores ou iguais à zero.\nPor favor, corrija-os e tente novamente!\n\nツ",
                                 "info")
        return validar_pesos

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
                self.exibir_mensagem(self.titulo_janela,
                                     f"PESO ZERO ENCONTRADO\n\nCertifique-se de que TODOS os pesos dos itens com unidade de medida em 'kg' (kilograma) sejam MAIORES QUE ZERO.\n\n{lista_codigos_peso_zero}\n\nPor favor, corrija-o(s) e tente novamente!\n\nツ",
                                     "info")
                return False
            else:
                return True

        except Exception as ex:
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}",
                                             "Erro ao validar pesos em unidade KG", 16 | 0)
            return False

    def validar_formato_campo_dimensao(self, dimensao):

        dimensao_sem_espaco = dimensao.replace(' ', '')

        if re.match(self.regex_campo_dimensao, dimensao_sem_espaco):
            return True
        else:
            return False

    def extrair_unidade_medida(self, valor_dimensao):
        padrao = r'[0-9.]+\s*(m³|m²|m)'
        unidade_extraida = re.findall(padrao, valor_dimensao, re.IGNORECASE)
        return unidade_extraida[0]

    def formatar_campos_dimensao(self, dataframe):
        items_mt_m2_dimensao_incorreta = {}
        items_unidade_incorreta = {}
        unidade_de_medida_totvs = ''
        df_campo_dimensao_formatado = dataframe.copy()

        for i, dimensao in enumerate(df_campo_dimensao_formatado.iloc[:, self.indice_coluna_dimensao]):
            codigo_filho = df_campo_dimensao_formatado.iloc[i, self.indice_coluna_codigo_excel]
            descricao = df_campo_dimensao_formatado.iloc[i, self.indice_coluna_descricao_excel]
            unidade_de_medida = self.obter_unidade_medida_codigo_filho(codigo_filho)

            if unidade_de_medida in ('MT', 'M2', 'M3') and self.validar_formato_campo_dimensao(str(dimensao)):
                unidade_de_medida_totvs = unidade_de_medida.replace('2', '²').replace('3', '³').replace('T', '').lower()
                if unidade_de_medida_totvs != self.extrair_unidade_medida(str(dimensao).lower()):
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
            self.exibir_mensagem(self.titulo_janela, mensagem_fixa + mensagem, "info")
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
            self.exibir_mensagem(self.titulo_janela, mensagem_fixa + mensagem, "info")
        if items_mt_m2_dimensao_incorreta or items_unidade_incorreta:
            self.excluir_arquivo_excel_bom(self.excel_file_path)
            sys.exit()

        return df_campo_dimensao_formatado

    def validacao_codigo_bloqueado(self, dataframe):
        try:
            conn = pyodbc.connect(
                f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}')
            cursor = conn.cursor()

            codigos_bloqueados = {}

            for i, row in dataframe.iterrows():
                codigo = row.iloc[self.indice_coluna_codigo_excel]
                descricao = row.iloc[self.indice_coluna_descricao_excel]
                query_retorna_valor_campo_bloqueio = f"""
                SELECT B1_MSBLQL FROM {self.database}.dbo.SB1010 WHERE B1_COD = '{codigo}' AND B1_REVATU <> 'ZZZ' AND D_E_L_E_T_ <> '*';
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
                self.exibir_mensagem(self.titulo_janela, mensagem_fixa + mensagem, 'warning')
                return False
            else:
                return True

        except Exception as e:
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o banco de dados. Erro: {str(e)}",
                                             "Erro ao consultar campo BLOQUEIO (B1_MSBLQL) na tabela produtos SG1010 do TOTVS",
                                             16 | 0)
        finally:
            if 'conn' in locals():
                conn.close()

    def extrair_numero(self, texto):
        numeros = ''
        for char in texto:
            if char.isdigit():
                numeros += char
        return numeros

    def verificar_codigo_filho_esta_correto_com_nome_do_desenho(self, dataframe):
        codigos_diferentes = {}
        codigo_nome_desenho_verificado = ''

        for index, row in dataframe.iterrows():
            codigo_filho = row.iloc[self.indice_coluna_codigo_excel].strip()
            codigo_nome_desenho = str(row.iloc[self.indice_coluna_nome_arquivo])
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
                codigo_filho_formatado = self.extrair_numero(codigo_filho).lstrip("0")
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
            mensagem_fixa = f"ATENÇÃO\n\n O CÓDIGO FILHO PODE ESTAR ERRADO.\n\n> APÓS A CORREÇÃO ATUALIZE O TEMPLATE DA BOM <\nO nome do desenho deve ser igual ao campo nº da peça.\n\nVerificar:\n\n"
            mensagem_final = "\nSMARTPLIC®"
            i = 1
            for code, description in codigos_diferentes.items():
                if code == 'nan':
                    mensagem_codigos += f"{i}. ALTERAR CONFIG. NOME USUÁRIO - {description[:20] + '...' if len(description) > 20 else description}\n\n"
                else:
                    mensagem_codigos += f"{i}. {code[:18] + '...' if len(code) > 18 else code} - {description[:20] + '...' if len(description) > 20 else description}\n\n"
                i += 1
            self.exibir_mensagem(self.titulo_janela, mensagem_fixa + mensagem_codigos + mensagem_final, "info")
            return False
        else:
            return True

    def verificar_se_template_bom_esta_correto(self, dataframe):
        if dataframe.shape[0] >= 1 and dataframe.shape[1] > 9:
            return True, "montagem"
        elif dataframe.shape[0] >= 1 and dataframe.shape[1] == 9:
            return True, "peca"
        elif dataframe.shape[0] == 0:
            self.exibir_mensagem(self.titulo_janela, f"ATENÇÃO!\n\nBOM VAZIA!\n\nツ\n\nSMARTPLIC®", "info")
            return False, ""
        else:
            self.exibir_mensagem(self.titulo_janela,
                                 f"ATENÇÃO!\n\nO TEMPLATE DA BOM FOI ATUALIZADO\n\nAtualize-o e tente novamente.\n\nツ\n\nSMARTPLIC®",
                                 "info")
            return False, ""

    def validacao_de_dados_bom(self, excel_file_path):
        df_excel = pd.read_excel(excel_file_path, sheet_name='Planilha1', header=None)

        # Encontra o índice da linha que contém "Pos." na segunda coluna
        pos_index = df_excel[df_excel.iloc[:, 0] == "Pos."].index[0]

        # Mantém todas as linhas até o índice encontrado
        df_excel = df_excel.iloc[:pos_index]

        df_excel.loc[:, self.indice_coluna_codigo_excel] = df_excel.loc[:, self.indice_coluna_codigo_excel].apply(
            lambda x: x.strip())

        bom_esta_correta, tipo_da_bom = self.verificar_se_template_bom_esta_correto(df_excel)
        validar_codigos = self.validar_formato_codigos_filho(df_excel)
        validar_quantidades = self.validacao_quantidades(df_excel)
        validar_descricoes = self.validar_descricao(df_excel.iloc[:, self.indice_coluna_descricao_excel])
        validar_pesos = self.validacao_pesos(df_excel)
        validar_codigo_filho_diferente_codigo_pai = self.verificar_codigo_filho_diferente_codigo_pai(self.nome_desenho,
                                                                                                     df_excel)

        if bom_esta_correta:
            if tipo_da_bom == "montagem":
                validar_codigo_filho_com_nome_desenho = self.verificar_codigo_filho_esta_correto_com_nome_do_desenho(
                    df_excel)
            elif tipo_da_bom == "peca":
                validar_codigo_filho_com_nome_desenho = True
            else:
                validar_codigo_filho_com_nome_desenho = False
        else:
            validar_codigo_filho_com_nome_desenho = False

        if validar_codigo_filho_com_nome_desenho:
            if validar_codigos.all() and validar_descricoes.all() and validar_quantidades.all() and validar_codigo_filho_diferente_codigo_pai.all() and validar_pesos.all() and validar_codigo_filho_com_nome_desenho:
                codigos_filho_tem_cadastro = self.verificar_cadastro_codigo_filho(
                    df_excel.iloc[:, self.indice_coluna_codigo_excel].tolist())
                if codigos_filho_tem_cadastro:
                    df_excel_campo_dimensao_tratado = self.formatar_campos_dimensao(df_excel)
                    bom_excel_sem_duplicatas = self.remover_linhas_duplicadas_e_consolidar_quantidade(
                        df_excel_campo_dimensao_tratado)
                    bom_excel_sem_duplicatas.iloc[:, self.indice_coluna_codigo_excel] = bom_excel_sem_duplicatas.iloc[:,
                                                                                        self.indice_coluna_codigo_excel].str.strip()
                    existe_codigo_filho_repetido = self.verificar_codigo_repetido(bom_excel_sem_duplicatas)
                    nao_existe_codigo_bloqueado = self.validacao_codigo_bloqueado(df_excel)
                    codigos_filho_tem_estrutura = self.verificar_se_existe_estrutura_codigos_filho(
                        bom_excel_sem_duplicatas.iloc[:, self.indice_coluna_codigo_excel].tolist())
                    pesos_maiores_que_zero_kg = self.validacao_pesos_unidade_kg(df_excel)
                    if codigos_filho_tem_cadastro and not existe_codigo_filho_repetido and nao_existe_codigo_bloqueado and codigos_filho_tem_estrutura and pesos_maiores_que_zero_kg:
                        return bom_excel_sem_duplicatas
        self.excluir_arquivo_excel_bom(excel_file_path)
        sys.exit()

    def atualizar_campo_revisao_do_codigo_pai(self, codigo_pai, numero_revisao):
        query_atualizar_campo_revisao = f"""UPDATE {self.database}.dbo.SB1010 SET B1_REVATU = N'{numero_revisao}' WHERE B1_COD = N'{codigo_pai}' AND D_E_L_E_T_ <> '*';"""
        try:
            # Uso do Context Manager para garantir o fechamento adequado da conexão
            with pyodbc.connect(
                    f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}') as conn:
                cursor = conn.cursor()
                cursor.execute(query_atualizar_campo_revisao)
                return True

        except Exception as ex:
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}",
                                             "Erro ao atualizar o campo REV. do código pai", 16 | 0)
            return False

    def atualizar_campo_data_ultima_revisao_do_codigo_pai(self, codigo_pai):
        data_atual = self.formatar_data_atual()
        query_atualizar_data_ultima_revisao = f"""UPDATE {self.database}.dbo.SB1010 SET B1_UREV = N'{data_atual}' WHERE B1_COD = N'{codigo_pai}' AND D_E_L_E_T_ <> '*';"""
        try:
            with pyodbc.connect(
                    f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}') as conn:
                cursor = conn.cursor()
                cursor.execute(query_atualizar_data_ultima_revisao)
                return True
        except Exception as ex:
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}",
                                             "Erro ao atualizar o campo DATA ULT. REV. do código pai", 16 | 0)
            return False

    def verificar_se_existe_estrutura_codigo_pai(self, codigo_pai):
        try:
            # Tente estabelecer a conexão com o banco de dados
            conn_str = f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}'
            engine = create_engine(f'mssql+pyodbc:///?odbc_connect={conn_str}')

            query_consulta_estrutura_totvs = f"""
                SELECT * FROM {self.database}.dbo.SG1010
                WHERE G1_COD = '{codigo_pai}' AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*'
                    AND G1_REVFIM = (
                        SELECT MAX(G1_REVFIM)
                        FROM {self.database}.dbo.SG1010
                        WHERE G1_COD = '{codigo_pai}' AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*'
                    );
            """

            # Executa a query SELECT e obtém os resultados em um DataFrame
            resultado_query_consulta_estrutura_totvs = pd.read_sql(query_consulta_estrutura_totvs, engine)

            return resultado_query_consulta_estrutura_totvs

        except Exception as ex:
            # Exibe uma caixa de diálogo se a conexão ou a consulta falhar
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}",
                                             "Erro ao verificar se existe estrutura no TOTVS", 16 | 0)

        finally:
            # Fecha a conexão com o banco de dados se estiver aberta
            if 'engine' in locals():
                engine.dispose()

    def obter_ultima_pk_tabela_estrutura(self):
        query_ultima_pk_tabela_estrutura = f"""SELECT TOP 1 R_E_C_N_O_ FROM {self.database}.dbo.SG1010 ORDER BY R_E_C_N_O_ DESC;"""
        try:
            # Uso do Context Manager para garantir o fechamento adequado da conexão
            with pyodbc.connect(
                    f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}') as conn:
                cursor = conn.cursor()
                cursor.execute(query_ultima_pk_tabela_estrutura)
                resultado_ultima_pk_tabela_estrutura = cursor.fetchone()
                valor_ultima_pk = resultado_ultima_pk_tabela_estrutura[0]

                return valor_ultima_pk

        except Exception as ex:
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}",
                                             "Erro ao obter última PK da tabela estrutura", 16 | 0)
            return None

    def obter_revisao_codigo_pai(self, codigo_pai, primeiro_cadastro):
        query_revisao_inicial = f"""SELECT B1_REVATU FROM {self.database}.dbo.SB1010 WHERE B1_COD = '{codigo_pai}'"""
        try:
            # Uso do Context Manager para garantir o fechamento adequado da conexão
            with pyodbc.connect(
                    f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}') as conn:
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
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}",
                                             "Erro ao consultar revisão do código pai na tabela de produtos", 16 | 0)
            return None

    def obter_unidade_medida_codigo_filho(self, codigo_filho):
        query_unidade_medida_codigo_filho = f"""SELECT B1_UM FROM {self.database}.dbo.SB1010 WHERE B1_COD = '{codigo_filho}'"""
        try:
            with pyodbc.connect(
                    f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}') as conn:
                cursor = conn.cursor()
                cursor.execute(query_unidade_medida_codigo_filho)

                unidade_medida = cursor.fetchone()
                valor_unidade_medida = unidade_medida[0]

                return valor_unidade_medida

        except Exception as ex:
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}",
                                             "Erro ao consultar Unidade de Medida de código filho", 16 | 0)
            return None

    def formatar_data_atual(self):
        # Formato yyyymmdd
        data_atual_formatada = date.today().strftime("%Y%m%d")
        return data_atual_formatada

    def verificar_cadastro_codigo_pai(self, codigo_pai):
        query_consulta_produto_codigo_pai = f"""SELECT B1_COD FROM {self.database}.dbo.SB1010 WHERE B1_COD = '{codigo_pai}'"""
        try:
            with pyodbc.connect(
                    f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}') as conn:
                cursor = conn.cursor()
                cursor.execute(query_consulta_produto_codigo_pai)

                resultado = cursor.fetchone()

                if resultado:
                    return True
                else:
                    self.exibir_mensagem(self.titulo_janela,
                                         f"O cadastro do item pai não foi encontrado!\n\nEfetue o cadastro do código {codigo_pai} e, em seguida, tente novamente.\n\nツ",
                                         "warning")
                    return False

        except Exception as ex:
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}",
                                             "Erro ao consultar cadastro do código pai", 16 | 0)
            return None

    def criar_nova_estrutura_totvs(self, codigo_pai, bom_excel_sem_duplicatas):
        primeiro_cadastro = True
        ultima_pk_tabela_estrutura = self.obter_ultima_pk_tabela_estrutura()
        revisao_inicial = self.obter_revisao_codigo_pai(codigo_pai, primeiro_cadastro)
        data_atual_formatada = self.formatar_data_atual()

        conn = pyodbc.connect(
            f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}')

        try:

            cursor = conn.cursor()

            for index, row in bom_excel_sem_duplicatas.iterrows():
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
                    (G1_FILIAL, G1_COD, G1_COMP, G1_TRT, G1_XUM, G1_QUANT, G1_PERDA, G1_INI, G1_FIM, G1_OBSERV, G1_FIXVAR, G1_GROPC, G1_OPC, G1_REVINI, G1_NIV, G1_NIVINV, G1_REVFIM, 
                    G1_OK, G1_POTENCI, G1_TIPVEC, G1_VECTOR, G1_VLCOMPE, G1_LOCCONS, G1_USAALT, G1_FANTASM, G1_LISTA, D_E_L_E_T_, R_E_C_N_O_, R_E_C_D_E_L_) 
                    VALUES (N'0101', N'{codigo_pai}  ', N'{codigo_filho}  ', N'   ', N'{unidade_medida}', {quantidade_formatada}, 0.0, N'{data_atual_formatada}', N'20491231', 
                    N'                                             ', N'V', N'   ', N'    ', N'{revisao_inicial}', N'01', N'99', N'{revisao_inicial}', N'    ', 0.0, N'      ', N'      ', 
                    N'N', N'  ', N'1', N' ', N'          ', 
                    N' ', {ultima_pk_tabela_estrutura}, 0);
                """

                cursor.execute(query_criar_nova_estrutura_totvs)

            conn.commit()

            self.exibir_mensagem(self.titulo_janela,
                                 f"Estrutura cadastrada com sucesso!\n\n{codigo_pai}\n\n( ͡° ͜ʖ ͡°)\n\nSMARTPLIC®",
                                 "info")
            return True, revisao_inicial

        except Exception as ex:
            ctypes.windll.user32.MessageBoxW(0,
                                             f"Falha na conexão ou consulta. Erro: {str(ex)} - PK-{ultima_pk_tabela_estrutura} - {codigo_pai} - {codigo_filho} - {quantidade} - {unidade_medida}",
                                             "Erro ao Criar Nova Estrutura", 16 | 0)
            return False

        finally:
            cursor.close()
            conn.close()

    def exibir_janela_mensagem_opcao(self, titulo, mensagem):
        root = tk.Tk()
        root.withdraw()  # Esconde a janela principal
        root.attributes('-topmost', True)  # Garante que a janela estará sempre no topo
        root.lift()  # Traz a janela para frente
        root.focus_force()  # Força o foco na janela

        # Mostrar a mensagem
        user_choice = messagebox.askquestion(
            titulo,
            mensagem,
            parent=root  # Define a janela principal como pai da mensagem
        )

        # Fechar a janela principal após a escolha do usuário
        root.destroy()

        if user_choice == "yes":
            return True
        else:
            return False

    def atualizar_itens_estrutura_totvs(self, codigo_pai, dataframe_codigos_em_comum):
        conn = pyodbc.connect(
            f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}')

        try:

            cursor = conn.cursor()

            for index, row in dataframe_codigos_em_comum.iterrows():
                codigo_filho = row.iloc[self.indice_coluna_codigo_excel]
                quantidade = row.iloc[self.indice_coluna_quantidade_excel]
                unidade_medida = self.obter_unidade_medida_codigo_filho(codigo_filho)

                if unidade_medida == 'KG':
                    quantidade = row.iloc[self.indice_coluna_peso_excel]
                elif unidade_medida in ('MT', 'M2', 'M3'):
                    quantidade = row.iloc[self.indice_coluna_dimensao]

                quantidade_formatada = "{:.2f}".format(float(quantidade))

                query_alterar_quantidade_estrutura = f"""UPDATE {self.database}.dbo.SG1010 SET G1_QUANT = {quantidade_formatada} WHERE G1_COD = '{codigo_pai}' AND G1_COMP = '{codigo_filho}'
                    AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*'
                    AND G1_REVFIM = (SELECT MAX(G1_REVFIM) FROM {self.database}.dbo.SG1010 WHERE G1_COD = '{codigo_pai}' AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*');
                """

                cursor.execute(query_alterar_quantidade_estrutura)

            conn.commit()

        except Exception as ex:
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou consulta. Erro: {str(ex)}",
                                             "Erro ao atualizar itens já existentes da estrutura", 16 | 0)

        finally:
            cursor.close()
            conn.close()

    def inserir_itens_estrutura_totvs(self, codigo_pai, dataframe_codigos_adicionados_bom,
                                      revisao_atualizada_estrutura):
        ultima_pk_tabela_estrutura = self.obter_ultima_pk_tabela_estrutura()
        data_atual_formatada = self.formatar_data_atual()
        revisao_inicial = revisao_atualizada_estrutura
        conn = pyodbc.connect(
            f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}')

        try:

            cursor = conn.cursor()

            for index, row in dataframe_codigos_adicionados_bom.iterrows():
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
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou consulta. Erro: {str(ex)}",
                                             "Erro ao inserir novos itens na estrutura", 16 | 0)
            return False

        finally:
            cursor.close()
            conn.close()

    def remover_itens_estrutura_totvs(self, codigo_pai, codigos_removidos_bom_df, revisao_anterior):
        conn = pyodbc.connect(
            f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}')

        try:

            cursor = conn.cursor()

            for index, row in codigos_removidos_bom_df.iterrows():
                codigo_filho = row.iloc[2]

                query_remover_itens_estrutura_totvs = f"""
                UPDATE {self.database}.dbo.SG1010
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
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou consulta. Erro: {str(ex)}",
                                             "Erro ao remover item da estrutura", 16 | 0)
            return False

        finally:
            cursor.close()
            conn.close()

    def resultado_comparacao(self):
        if self.codigos_em_comum:
            ctypes.windll.user32.MessageBoxW(
                0, f"Códigos em comuns: {self.codigos_em_comum}", "ITENS EM COMUM", 1)

        if self.codigos_adicionados_bom:
            ctypes.windll.user32.MessageBoxW(
                0, f"Itens adicionados: {self.codigos_adicionados_bom}", "ITENS ADICIONADOS", 1)

        if self.codigos_removidos_bom:
            ctypes.windll.user32.MessageBoxW(
                0, f"Itens removidos: {self.codigos_removidos_bom}", "ITENS REMOVIDOS", 1)

    def comparar_bom_com_totvs(self, bom_excel_sem_duplicatas, resultado_query_consulta_estrutura_totvs):
        resultado_query_consulta_estrutura_totvs['G1_COMP'] = resultado_query_consulta_estrutura_totvs[
            'G1_COMP'].str.strip()

        # Códigos em comum
        self.codigos_em_comum_df = bom_excel_sem_duplicatas[
            bom_excel_sem_duplicatas.iloc[:, self.indice_coluna_codigo_excel].isin(
                resultado_query_consulta_estrutura_totvs['G1_COMP'])].copy()

        # Códigos adicionados no BOM
        self.codigos_adicionados_bom_df = bom_excel_sem_duplicatas[
            ~bom_excel_sem_duplicatas.iloc[:, self.indice_coluna_codigo_excel].isin(
                resultado_query_consulta_estrutura_totvs['G1_COMP'])].copy()

        # Códigos removidos no BOM
        self.codigos_removidos_bom_df = resultado_query_consulta_estrutura_totvs[
            ~resultado_query_consulta_estrutura_totvs['G1_COMP'].isin(
                bom_excel_sem_duplicatas.iloc[:, self.indice_coluna_codigo_excel])].copy()

        return self.codigos_em_comum_df, self.codigos_adicionados_bom_df, self.codigos_removidos_bom_df

        #resultado_comparacao()

    def atualizar_campo_revfim_codigos_existentes(self, codigo_pai, revisao_anterior, revisao_atualizada):
        conn = pyodbc.connect(
            f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}')
        try:
            cursor = conn.cursor()

            for index, row in self.codigos_em_comum_df.iterrows():
                codigo_filho = row.iloc[self.indice_coluna_codigo_excel]

                query_atualizar_campo_revfim_estrutura = f"""UPDATE {self.database}.dbo.SG1010 SET G1_REVFIM = N'{revisao_atualizada}' WHERE G1_COD = '{codigo_pai}' AND G1_COMP = '{codigo_filho}'
                    AND G1_REVFIM = N'{revisao_anterior}' AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*'
                """

                cursor.execute(query_atualizar_campo_revfim_estrutura)

            conn.commit()

            return True

        except Exception as ex:
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou consulta. Erro: {str(ex)}",
                                             "Erro ao atualizar campo revfim dos itens já existentes da estrutura",
                                             16 | 0)
            return False

        finally:
            cursor.close()
            conn.close()

    def calculo_revisao_anterior(self, revisao_atualizada):
        revisao_atualizada = int(revisao_atualizada) - 1
        revisao_anterior = str(revisao_atualizada).zfill(3)
        return revisao_anterior

    def exibir_mensagem(self, title, message, icon_type):
        root = tk.Tk()
        root.withdraw()
        root.lift()  # Garante que a janela esteja na frente
        root.title(title)
        root.attributes('-topmost', True)

        if icon_type == 'info':
            messagebox.showinfo(title, message)
        elif icon_type == 'warning':
            messagebox.showwarning(title, message)
        elif icon_type == 'error':
            messagebox.showerror(title, message)

        root.destroy()

    def start_task(self):
        thread = threading.Thread(target=self.executar_logica)
        thread.start()

    def update_progress(self, value):
        self.progress['value'] = value
        root.update_idletasks()

    def executar_logica(self):
        delay = 0.4
        self.status_label.config(text="Iniciando cadastro...")
        time.sleep(0.7)
        excel_file_path = self.obter_caminho_arquivo_excel(self.nome_desenho)
        self.update_progress(10)

        self.status_label.config(text="Validando formato do código pai...")
        time.sleep(delay)
        formato_codigo_pai_correto = self.validar_formato_codigo_pai(self.nome_desenho)
        nova_estrutura_cadastrada = False
        self.update_progress(20)

        existe_cadastro_codigo_pai = False
        if formato_codigo_pai_correto:
            self.status_label.config(text="Verificando cadastro do código pai...")
            time.sleep(delay)
            existe_cadastro_codigo_pai = self.verificar_cadastro_codigo_pai(self.nome_desenho)
            self.status_label.config(text="")
            self.update_progress(30)

        if formato_codigo_pai_correto and existe_cadastro_codigo_pai:
            self.status_label.config(text="Validando dados da tabela de BOM...")
            time.sleep(delay)
            bom_excel_sem_duplicatas = self.validacao_de_dados_bom(excel_file_path)
            self.update_progress(40)
            self.status_label.config(text="Verificando se já existe estrutura cadastrada...")
            time.sleep(delay)
            resultado_estrutura_codigo_pai = self.verificar_se_existe_estrutura_codigo_pai(self.nome_desenho)
            self.update_progress(50)

            if not bom_excel_sem_duplicatas.empty and resultado_estrutura_codigo_pai.empty:
                self.status_label.config(text="Cadastrando estrutura...")
                time.sleep(delay)
                nova_estrutura_cadastrada, revisao_atualizada = self.criar_nova_estrutura_totvs(self.nome_desenho,
                                                                                                bom_excel_sem_duplicatas)
                self.status_label.config(text="Atualizando revisão da estrutura...")
                time.sleep(delay)
                self.atualizar_campo_revisao_do_codigo_pai(self.nome_desenho, revisao_atualizada)
                self.update_progress(60)

            if bom_excel_sem_duplicatas.empty and not nova_estrutura_cadastrada:
                self.exibir_mensagem(self.titulo_janela,
                                     f"OPS!\n\nA BOM está vazia!\n\nPor gentileza, preencha adequadamente a BOM e tente novamente!\n\n{self.nome_desenho}\n\nツ\n\nSMARTPLIC®",
                                     "warning")
                self.update_progress(100)
            if not bom_excel_sem_duplicatas.empty and not resultado_estrutura_codigo_pai.empty:
                mensagem = f"ESTRUTURA EXISTENTE\n\nJá existe uma estrutura cadastrada no TOTVS para este produto!\n\n{self.nome_desenho}\n\nDeseja realizar a alteração da estrutura?"
                usuario_quer_alterar = self.exibir_janela_mensagem_opcao(self.titulo_janela, mensagem)
                self.update_progress(70)

                if usuario_quer_alterar:
                    resultado = self.comparar_bom_com_totvs(bom_excel_sem_duplicatas, resultado_estrutura_codigo_pai)
                    codigos_em_comum_df, codigos_adicionados_bom_df, codigos_removidos_bom_df = resultado
                    self.update_progress(80)

                    if not codigos_em_comum_df.empty:
                        self.status_label.config(text="Atualizando as quantidades da estrutura...")
                        time.sleep(delay)
                        self.atualizar_itens_estrutura_totvs(self.nome_desenho, codigos_em_comum_df)
                        self.update_progress(100)

                    if not codigos_adicionados_bom_df.empty or not codigos_removidos_bom_df.empty:
                        primeiro_cadastro = False
                        revisao_atualizada = self.obter_revisao_codigo_pai(self.nome_desenho, primeiro_cadastro)
                        itens_adicionados_sucesso = False
                        itens_removidos_sucesso = False
                        revisao_anterior = self.calculo_revisao_anterior(revisao_atualizada)
                        self.update_progress(90)

                        if not codigos_adicionados_bom_df.empty:
                            self.status_label.config(text="Inserindo itens na estrutura...")
                            time.sleep(delay)
                            itens_adicionados_sucesso = self.inserir_itens_estrutura_totvs(self.nome_desenho,
                                                                                           codigos_adicionados_bom_df,
                                                                                           revisao_atualizada)
                            self.update_progress(95)

                        if not codigos_removidos_bom_df.empty:
                            self.status_label.config(text="Removendo itens da estrutura...")
                            time.sleep(delay)
                            itens_removidos_sucesso = self.remover_itens_estrutura_totvs(self.nome_desenho,
                                                                                         codigos_removidos_bom_df,
                                                                                         revisao_anterior)
                            self.update_progress(98)

                        if itens_adicionados_sucesso or itens_removidos_sucesso:
                            self.status_label.config(text="Atualizando revisão da estrutura...")
                            time.sleep(delay)
                            self.atualizar_campo_revfim_codigos_existentes(self.nome_desenho, revisao_anterior,
                                                                           revisao_atualizada)
                            self.atualizar_campo_revisao_do_codigo_pai(self.nome_desenho, revisao_atualizada)
                            self.atualizar_campo_data_ultima_revisao_do_codigo_pai(self.nome_desenho)
                            self.update_progress(100)
                            self.status_label.config(text="Alteração da estrutura finalizada!")
                            time.sleep(delay)
                            self.exibir_mensagem(self.titulo_janela,
                                                 f"Alteração da estrutura realizada com sucesso!\n\n{self.nome_desenho}\n\n( ͡° ͜ʖ ͡°)\n\nSMARTPLIC®",
                                                 "info")

                    else:
                        self.status_label.config(text="Atualização de quantidades finalizada!")
                        time.sleep(delay)
                        self.exibir_mensagem(self.titulo_janela,
                                             f"Quantidades atualizadas com sucesso!\n\nNão foi adicionado e/ou removido itens da estrutura.\n\n{self.nome_desenho}\n\n( ͡° ͜ʖ ͡°)\n\nSMARTPLIC®",
                                             "info")
        end_time = time.time()
        elapsed = end_time - self.start_time
        self.status_label.config(text=f"Processo finalizado!\n\n{elapsed:.3f} segundos\n\nEUREKA®")
        self.update_progress(100)
        self.excluir_arquivo_excel_bom(excel_file_path)


if __name__ == "__main__":
    root = tk.Tk()
    cadastro = CadastrarBomTOTVS(root)
    cadastro.start_task()
    root.attributes('-topmost', True)
    root.geometry("400x200")
    root.mainloop()
