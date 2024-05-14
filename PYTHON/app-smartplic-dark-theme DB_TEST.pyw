import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QHBoxLayout, QTableWidget, \
    QTableWidgetItem, QHeaderView, QSizePolicy, QSpacerItem, QMessageBox, QFileDialog, QToolButton, QTabWidget, QItemDelegate, QAbstractItemView, QCheckBox
from PyQt5.QtGui import QFont, QIcon, QDesktopServices, QColor
from PyQt5.QtCore import Qt, QUrl, QCoreApplication, pyqtSignal
import pyodbc
import pyperclip
import os
import time
import pandas as pd
import ctypes
from datetime import datetime
import tkinter as tk
from tkinter import messagebox

class ConsultaApp(QWidget):
    
    # Adicione este sinal à classe
    guia_fechada = pyqtSignal()
        
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("SMARTPLIC® v2.2.1 - Dark theme - DBTEST")
        
        # Configurar o ícone da janela
        icon_path = "010.png"
        self.setWindowIcon(QIcon(icon_path))
        
        # Ajuste a cor de fundo da janela
        self.setAutoFillBackground(True)
        palette = self.palette()
        palette.setColor(self.backgroundRole(), QColor('#363636'))  # Substitua pela cor desejada
        self.setPalette(palette)
            
        self.nova_janela = None  # Adicione esta linha
        
        self.tabWidget = QTabWidget(self)  # Adicione um QTabWidget ao layout principal
        self.tabWidget.setTabsClosable(True)  # Adicione essa linha para permitir o fechamento de guias
        self.tabWidget.tabCloseRequested.connect(self.fechar_guia)
        self.tabWidget.setVisible(False)  # Inicialmente, a guia está invisível

        # Aplicar folha de estilo ao aplicativo
        self.setStyleSheet("""
            * {
                background-color: #363636;
            }
            
            QLabel, QCheckBox {
                color: #EEEEEE;
                font-size: 11px;
                font-weight: bold;
            }

            QLineEdit {
                background-color: #A7A6A6;
                border: 1px solid #262626;
                padding: 5px;
                border-radius: 8px;
            }

            QPushButton {
                background-color: #0a79f8;
                color: #fff;
                padding: 5px 15px;
                border: 2px;
                border-radius: 8px;
                font-size: 11px;
                height: 20px;
                font-weight: bold;
                margin-top: 6px;
                margin-bottom: 6px;
            }

            QPushButton:hover {
                background-color: #fff;
                color: #0a79f8
            }

            QPushButton:pressed {
                background-color: #6703c5;
                color: #fff;
            }

            QTableWidget {
                border: 1px solid #000000;
                background-color: #363636;
            }

            QTableWidget QHeaderView::section {
                background-color: #262626;
                color: #A7A6A6;
                padding: 5px;
                height: 18px;
            }

            QTableWidget QHeaderView::section:horizontal {
                border-top: 1px solid #333;
            }
            
            QTableWidget::item {
                background-color: #363636;
                color: #fff;
                font-weight: bold;
            }
            
            QTableWidget::item:selected {
                background-color: #000000;
                color: #EEEEEE;
                font-weight: bold;
            }
        """)
            
        self.codigo_var = QLineEdit(self)
        self.descricao_var = QLineEdit(self)
        self.descricao2_var = QLineEdit(self)
        self.tipo_var = QLineEdit(self)
        self.um_var = QLineEdit(self)
        self.armazem_var = QLineEdit(self)
        self.grupo_var = QLineEdit(self)
        self.grupo_desc_var = QLineEdit(self)
        
        fonte = "Segoe UI"
        tamanho_fonte = 10

        self.codigo_var.setFont(QFont(fonte, tamanho_fonte))  # Substitua "Arial" pela fonte desejada e 12 pelo tamanho desejado
        self.descricao_var.setFont(QFont(fonte, tamanho_fonte))
        self.descricao2_var.setFont(QFont(fonte, tamanho_fonte))
        self.tipo_var.setFont(QFont(fonte, tamanho_fonte))
        self.um_var.setFont(QFont(fonte, tamanho_fonte))
        self.armazem_var.setFont(QFont(fonte, tamanho_fonte))
        self.grupo_var.setFont(QFont(fonte, tamanho_fonte))
        self.grupo_desc_var.setFont(QFont(fonte, tamanho_fonte))

        self.configurar_campos()

        self.btn_consultar = QPushButton("Pesquisar", self)
        self.btn_consultar.clicked.connect(self.executar_consulta)
        self.btn_consultar.setMinimumWidth(100)
        
        self.btn_consultar_estrutura = QPushButton("Consultar Estrutura", self)
        self.btn_consultar_estrutura.clicked.connect(self.executar_consulta_estrutura)
        self.btn_consultar_estrutura.setMinimumWidth(150)
        self.btn_consultar_estrutura.setEnabled(False)
        
        self.btn_onde_e_usado = QPushButton("Onde é usado?", self)
        self.btn_onde_e_usado.clicked.connect(self.executar_consulta_onde_usado)
        self.btn_onde_e_usado.setMinimumWidth(150)
        self.btn_onde_e_usado.setEnabled(False)
        
        self.btn_limpar = QPushButton("Limpar", self)
        self.btn_limpar.clicked.connect(self.limpar_campos)
        self.btn_limpar.setMinimumWidth(100)
        
        self.btn_nova_janela = QPushButton("Nova Janela", self)
        self.btn_nova_janela.clicked.connect(self.abrir_nova_janela)
        self.btn_nova_janela.setMinimumWidth(100)
        
        self.btn_abrir_desenho = QPushButton("Abrir Desenho", self)
        self.btn_abrir_desenho.clicked.connect(self.abrir_desenho)
        self.btn_abrir_desenho.setMinimumWidth(100)
        
        self.btn_exportar_excel = QPushButton("Exportar Excel", self)
        self.btn_exportar_excel.clicked.connect(self.exportar_excel)
        self.btn_exportar_excel.setMinimumWidth(100)
        self.btn_exportar_excel.setEnabled(False)  # Desativar inicialmente
        
        self.btn_calculo_peso = QPushButton("Tabela de pesos", self)
        self.btn_calculo_peso.clicked.connect(self.abrir_tabela_pesos)
        self.btn_calculo_peso.setMinimumWidth(100)
        
        self.btn_fechar = QPushButton("Fechar", self)
        self.btn_fechar.clicked.connect(self.fechar_janela)
        self.btn_fechar.setMinimumWidth(100)

        self.configurar_tabela()
        
        # Configurar a tabela
        self.configurar_tabela_tooltips()

        # Conectar o evento returnPressed dos campos de entrada ao método executar_consulta
        self.codigo_var.returnPressed.connect(self.executar_consulta)
        self.descricao_var.returnPressed.connect(self.executar_consulta)
        self.descricao2_var.returnPressed.connect(self.executar_consulta)
        self.tipo_var.returnPressed.connect(self.executar_consulta)
        self.um_var.returnPressed.connect(self.executar_consulta)
        self.armazem_var.returnPressed.connect(self.executar_consulta)
        self.grupo_var.returnPressed.connect(self.executar_consulta)
        self.grupo_desc_var.returnPressed.connect(self.executar_consulta)

        layout = QVBoxLayout()
        layout_linha_01 = QHBoxLayout()
        layout_linha_02 = QHBoxLayout()
        layout_linha_03 = QHBoxLayout()

        layout_linha_01.addWidget(QLabel("Código:"))
        layout_linha_01.addWidget(self.codigo_var)
        layout_linha_01.addWidget(self.criar_botao_limpar(self.codigo_var))

        layout_linha_01.addWidget(QLabel("Descrição:"))
        layout_linha_01.addWidget(self.descricao_var)
        layout_linha_01.addWidget(self.criar_botao_limpar(self.descricao_var))

        layout_linha_01.addWidget(QLabel("Contém na Descrição:"))
        layout_linha_01.addWidget(self.descricao2_var)
        layout_linha_01.addWidget(self.criar_botao_limpar(self.descricao2_var))

        layout_linha_02.addWidget(QLabel("Tipo:"))
        layout_linha_02.addWidget(self.tipo_var)
        layout_linha_02.addWidget(self.criar_botao_limpar(self.tipo_var))

        layout_linha_02.addWidget(QLabel("Unid. Medida:"))
        layout_linha_02.addWidget(self.um_var)
        layout_linha_02.addWidget(self.criar_botao_limpar(self.um_var))

        layout_linha_02.addWidget(QLabel("Armazém:"))
        layout_linha_02.addWidget(self.armazem_var)
        layout_linha_02.addWidget(self.criar_botao_limpar(self.armazem_var))

        layout_linha_02.addWidget(QLabel("Grupo:"))
        layout_linha_02.addWidget(self.grupo_var)
        layout_linha_02.addWidget(self.criar_botao_limpar(self.grupo_var))
        
        layout_linha_02.addWidget(QLabel("Desc. Grupo:"))
        layout_linha_02.addWidget(self.grupo_desc_var)
        layout_linha_02.addWidget(self.criar_botao_limpar(self.grupo_desc_var))
        
        self.checkbox_bloqueado = QCheckBox("Bloqueado?", self)
        layout_linha_02.addWidget(self.checkbox_bloqueado)
        
        layout_linha_03.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        
        layout_linha_03.addWidget(self.btn_consultar)
        layout_linha_03.addWidget(self.btn_consultar_estrutura)
        layout_linha_03.addWidget(self.btn_onde_e_usado)
        layout_linha_03.addWidget(self.btn_limpar)
        layout_linha_03.addWidget(self.btn_nova_janela)
        layout_linha_03.addWidget(self.btn_abrir_desenho)
        layout_linha_03.addWidget(self.btn_exportar_excel)
        layout_linha_03.addWidget(self.btn_calculo_peso)
        layout_linha_03.addWidget(self.btn_fechar)
        
        # Adicione um espaçador esticável para centralizar os botões
        layout_linha_03.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        
        layout.addLayout(layout_linha_01)
        layout.addLayout(layout_linha_02)
        layout.addLayout(layout_linha_03)
                
        layout.addWidget(self.tree)
        
        layout.addWidget(self.tabWidget)  # Adicione o QTabWidget ao layout principal

        self.setLayout(layout)
        
        self.guias_abertas = []
        self.guias_abertas_onde_usado = []
    
    def criar_botao_limpar(self, campo):
        botao_limpar = QToolButton(self)
        botao_limpar.setIcon(QIcon('clear_icon.png'))  # Substitua 'icone_limpar.png' pelo caminho do ícone desejado
        botao_limpar.setCursor(Qt.PointingHandCursor)
        botao_limpar.clicked.connect(lambda: campo.clear())
        return botao_limpar

    
    def exportar_excel(self):
        # Obter o caminho do arquivo para salvar
        file_path, _ = QFileDialog.getSaveFileName(self, 'Salvar como', '', 'Arquivos Excel (*.xlsx);;Todos os arquivos (*)')

        if file_path:
            # Obter os dados da tabela
            data = self.obter_dados_tabela()

            # Criar um DataFrame pandas
            df = pd.DataFrame(data, columns=["CÓDIGO", "DESCRIÇÃO", "DESC. COMP.", "TIPO", "UM", "ARMAZÉM","GRUPO", "DESC. GRUPO", "CC", "BLOQUEADO?", "REV.", ""])

            # Salvar o DataFrame como um arquivo Excel
            df.to_excel(file_path, index=False)
            
    def obter_dados_tabela(self):
        # Obter os dados da tabela
        data = []
        for i in range(self.tree.rowCount()):
            row_data = []
            for j in range(self.tree.columnCount()):
                item = self.tree.item(i, j)
                if item is not None:
                    row_data.append(item.text())
                else:
                    row_data.append("")
            data.append(row_data)
        return data

    def configurar_campos(self):
        self.codigo_var
        self.descricao_var
        self.descricao2_var
        self.tipo_var
        self.um_var
        self.armazem_var
        self.grupo_var
        self.grupo_desc_var

    def configurar_tabela(self):
        self.tree = QTableWidget(self)
        self.tree.setColumnCount(12)
        self.tree.setHorizontalHeaderLabels(
            ["CÓDIGO", "DESCRIÇÃO", "DESC. COMP.", "TIPO", "UM", "ARMAZÉM", "GRUPO", "DESC. GRUPO", "CC", "BLOQUEADO?","REV.", ""])
        self.tree.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tree.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tree.setSelectionBehavior(QTableWidget.SelectRows)
        self.tree.setSelectionMode(QTableWidget.SingleSelection)
        
        # Conectar o evento itemDoubleClicked ao método copiar_linha
        self.tree.itemDoubleClicked.connect(self.copiar_linha)
        
        # Configurar a fonte da tabela
        fonte_tabela = QFont("Segoe UI", 10)  # Substitua por sua fonte desejada e tamanho
        self.tree.setFont(fonte_tabela)

        # Ajustar a altura das linhas
        altura_linha = 34  # Substitua pelo valor desejado
        self.tree.verticalHeader().setDefaultSectionSize(altura_linha)
        
        # Conectar o evento sectionClicked ao método ordenar_tabela
        self.tree.horizontalHeader().sectionClicked.connect(self.ordenar_tabela)
        
        # Redimensionar a última coluna para preencher o espaço restante
        self.tree.horizontalHeader().setStretchLastSection(True)
        
    def configurar_tabela_tooltips(self):
        headers = ["CÓDIGO", "DESCRIÇÃO", "DESC. COMP.", "TIPO", "UM", "ARMAZÉM", "GRUPO", "DESC. GRUPO", "CC", "BLOQUEADO?", "REV."]
        tooltips = [
            "Código do produto",
            "Descrição do produto",
            "Descrição completa do produto",
            "Tipo de produto\n\nMC - Material de consumo\nMP - Matéria-prima\nPA - Produto Acabado\nPI - Produto Intermediário\nSV - Serviço",
            "Unidade de medida",
            "Armazém padrão\n\n01 - Matéria-prima\n02 - Produto Intermediário\n03 - Produto Comercial\n04 - Produto Acabado",
            "Grupo do produto",
            "Descrição do grupo do produto",
            "Centro de custo",
            "Indica se o produto está bloqueado",
            "Revisão atual do produto"
        ]

        for i, header in enumerate(headers):
            self.tree.setHorizontalHeaderItem(i, QTableWidgetItem(header))
            self.tree.horizontalHeaderItem(i).setToolTip(tooltips[i])
        
    def copiar_linha(self, item):
        # Verificar se um item foi clicado
        if item is not None:
            valor_campo = item.text()
            pyperclip.copy(str(valor_campo))
        
    def ordenar_tabela(self, logicalIndex):
        # Obter o índice real da coluna (considerando a ordem de classificação)
        index = self.tree.horizontalHeader().sortIndicatorOrder()

        # Definir a ordem de classificação
        order = Qt.AscendingOrder if index == 0 else Qt.DescendingOrder

        # Ordenar a tabela pela coluna clicada
        self.tree.sortItems(logicalIndex, order)
        
    def limpar_campos(self):
        # Limpar os dados dos campos
        self.codigo_var.clear()
        self.descricao_var.clear()
        self.descricao2_var.clear()
        self.tipo_var.clear()
        self.um_var.clear()
        self.armazem_var.clear()
        self.grupo_var.clear()
        self.grupo_desc_var.clear()
        
    def bloquear_campos_pesquisa(self):
        # Bloquear campos de pesquisa
        self.codigo_var.setEnabled(False)
        self.descricao_var.setEnabled(False)
        self.descricao2_var.setEnabled(False)
        self.tipo_var.setEnabled(False)
        self.um_var.setEnabled(False)
        self.armazem_var.setEnabled(False)
        self.grupo_var.setEnabled(False)
        self.grupo_desc_var.setEnabled(False)
    
        # Desativar os botões após o carregamento da tabela
        self.btn_consultar.setEnabled(False)
        self.btn_exportar_excel.setEnabled(False)
        self.btn_consultar_estrutura.setEnabled(False)
        self.btn_onde_e_usado.setEnabled(False)
        
    def desbloquear_campos_pesquisa(self):
        # Desbloquear campos de pesquisa
        self.codigo_var.setEnabled(True)
        self.descricao_var.setEnabled(True)
        self.descricao2_var.setEnabled(True)
        self.tipo_var.setEnabled(True)
        self.um_var.setEnabled(True)
        self.armazem_var.setEnabled(True)
        self.grupo_var.setEnabled(True)
        self.grupo_desc_var.setEnabled(True)
    
        # Ativar o botões após o carregamento da tabela
        self.btn_consultar.setEnabled(True)
        self.btn_exportar_excel.setEnabled(True)
        self.btn_consultar_estrutura.setEnabled(True)
        self.btn_onde_e_usado.setEnabled(True)
        
    
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
        
    def selecionar_query_conforme_filtro(self):
        # Obter os valores dos campos de consulta
        codigo = self.codigo_var.text().upper().strip()
        descricao = self.descricao_var.text().upper().strip()
        descricao2 = self.descricao2_var.text().upper().strip()
        tipo = self.tipo_var.text().upper().strip()
        um = self.um_var.text().upper().strip()
        armazem = self.armazem_var.text().upper().strip()
        grupo = self.grupo_var.text().upper().strip()
        desc_grupo = self.grupo_desc_var.text().upper().strip()
        status_checkbox = self.checkbox_bloqueado.isChecked()
        
        if codigo == '' and descricao == '' and descricao2 == '' and tipo == '' and um == '' and armazem == '' and grupo == '' and desc_grupo == '':
            self.btn_consultar.setEnabled(False)
            self.exibir_mensagem("ATENÇÃO!", "Os campos de pesquisa estão vazios.\nPreencha algum campo e tente novamente.\n\nツ\n\nSMARTPLIC®", "info")
            return True
        
        if status_checkbox:
            status_bloqueado = '1'
            return f"""
                SELECT B1_COD, B1_DESC, B1_XDESC2, B1_TIPO, B1_UM, B1_LOCPAD, B1_GRUPO, B1_ZZNOGRP, B1_CC, B1_MSBLQL, B1_REVATU
                FROM {database}.dbo.SB1010
                WHERE B1_COD LIKE '{codigo}%' AND B1_DESC LIKE '{descricao}%' AND B1_DESC LIKE '%{descricao2}%'
                AND B1_TIPO LIKE '{tipo}%' AND B1_UM LIKE '{um}%' AND B1_LOCPAD LIKE '{armazem}%' AND B1_GRUPO LIKE '{grupo}%' 
                AND B1_ZZNOGRP LIKE '%{desc_grupo}%' AND B1_MSBLQL = '{status_bloqueado}'
                AND D_E_L_E_T_ <> '*'
                ORDER BY B1_COD ASC"""
        else:
            return f"""
                SELECT B1_COD, B1_DESC, B1_XDESC2, B1_TIPO, B1_UM, B1_LOCPAD, B1_GRUPO, B1_ZZNOGRP, B1_CC, B1_MSBLQL, B1_REVATU
                FROM {database}.dbo.SB1010
                WHERE B1_COD LIKE '{codigo}%' AND B1_DESC LIKE '{descricao}%' AND B1_DESC LIKE '%{descricao2}%'
                AND B1_TIPO LIKE '{tipo}%' AND B1_UM LIKE '{um}%' AND B1_LOCPAD LIKE '{armazem}%' AND B1_GRUPO LIKE '{grupo}%' 
                AND B1_ZZNOGRP LIKE '%{desc_grupo}%' AND D_E_L_E_T_ <> '*'
                ORDER BY B1_COD ASC"""


    def executar_consulta(self):
        
        select_query = self.selecionar_query_conforme_filtro()
        
        if isinstance(select_query, bool) and select_query:
            self.btn_consultar.setEnabled(True)
            return
        
        self.bloquear_campos_pesquisa()

        try:
            # Estabelecer a conexão com o banco de dados
            conn = pyodbc.connect(
                f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')

            # Criar um cursor para executar comandos SQL
            cursor = conn.cursor()

            # Executar a consulta
            cursor.execute(select_query)
            
            # Limpar a ordenação
            self.tree.horizontalHeader().setSortIndicator(-1, Qt.AscendingOrder)

            # Limpar a tabela
            self.tree.setRowCount(0)
            
            # Definir cores alternadas
            cores = [QColor("#FFFFFF"), QColor("#FFFFFF")]
            
            time.sleep(0.1)

            # Preencher a tabela com os resultados
            for i, row in enumerate(cursor.fetchall()):
                # Calcular índice da cor alternada
                indice_cor = i % 2
                cor_fundo = cores[indice_cor]

                self.tree.setSortingEnabled(False)  # Permitir ordenação
                # Inserir os valores formatados na tabela
                self.tree.insertRow(i)
                for j, value in enumerate(row):
                    if j == 9:  # Verifica se o valor é da coluna B1_MSBLQL
                        # Converte o valor 1 para 'Sim' e 2 para 'Não'
                        if value == '1':
                            value = 'Sim'
                        else:
                            value = 'Não'
                    item = QTableWidgetItem(str(value).strip())
                    item.setBackground(cor_fundo)  # Definir cor de fundo
                    self.tree.setItem(i, j, item)

                # Permitir que a interface gráfica seja atualizada
                QCoreApplication.processEvents()

            self.tree.setSortingEnabled(True)  # Permitir ordenação
            
            self.desbloquear_campos_pesquisa()
            
        except pyodbc.Error as ex:
            print(f"Falha na consulta. Erro: {str(ex)}")

        finally:
            # Fechar a conexão com o banco de dados
            conn.close()
            
            
    def abrir_desenho(self):
        item_selecionado = self.tree.currentItem()

        if item_selecionado:
            codigo = self.tree.item(item_selecionado.row(), 0).text()
            pdf_path = os.path.join("Y:/PDF-OFICIAL/", f"{codigo}.PDF")
            pdf_path = os.path.normpath(pdf_path)

            if os.path.exists(pdf_path):
                QCoreApplication.processEvents()
                QDesktopServices.openUrl(QUrl.fromLocalFile(pdf_path))
            else:
                mensagem = f"Desenho não encontrado!\n\n:-("
                QMessageBox.information(self, f"{codigo}", mensagem)

                
    def abrir_tabela_pesos(self):
        os.startfile(r'\\192.175.175.4\f\INTEGRANTES\ELIEZER\DOCUMENTOS_UTEIS\TABELA_PESO.xlsx')  
            
            
    def abrir_nova_janela(self):
        if not self.nova_janela or not self.nova_janela.isVisible():
            self.nova_janela = ConsultaApp()
            self.nova_janela.setGeometry(self.x() + 50, self.y() + 50, self.width(), self.height())
            self.nova_janela.show()
            
            
    def fechar_janela(self):
        self.close()
        
        
    def fechar_guia(self, index):
        if index >= 0:
            try:
                codigo_guia_fechada = self.tabWidget.tabText(index)
                self.guias_abertas.remove(codigo_guia_fechada)
            
            # Por ter duas listas de controle de abas abertas, 'guias_abertas = []' e 'guias_abertas_onde_usado = []',
            # ao fechar uma guia ocorre uma exceção (ValueError) se o código não for encontrado em uma das listas.
            # Utilize try/except para contornar esse problema.   
            except ValueError:
                codigo_guia_fechada = self.tabWidget.tabText(index).split(' - ')[1]
                self.guias_abertas_onde_usado.remove(codigo_guia_fechada)
                        
            finally:
                self.tabWidget.removeTab(index)

                if not self.existe_guias_abertas():
                    # Se não houver mais guias abertas, remova a guia do layout principal
                    self.tabWidget.setVisible(False)
                    self.guia_fechada.emit()
    
    
    def existe_guias_abertas(self):
        return self.tabWidget.count() > 0
    
    
    def ajustar_largura_coluna_descricao(self, tree_widget):
        header = tree_widget.horizontalHeader()
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)


    def executar_consulta_estrutura(self):
        item_selecionado = self.tree.currentItem()

        if item_selecionado:
            codigo = self.tree.item(item_selecionado.row(), 0).text()
            descricao_onde_usado = self.tree.item(item_selecionado.row(), 1).text()
            
            if codigo not in self.guias_abertas:
                select_query_estrutura = f"""
                    SELECT struct.G1_COMP AS CÓDIGO, prod.B1_DESC AS DESCRIÇÃO, struct.G1_QUANT AS "QTD.", struct.G1_XUM AS "UNID.", struct.G1_REVFIM AS "REVISÃO", 
                    struct.G1_INI AS "INSERIDO EM:"
                    FROM {database}.dbo.SG1010 struct
                    INNER JOIN {database}.dbo.SB1010 prod
                    ON struct.G1_COMP = prod.B1_COD
                    WHERE G1_COD = '{codigo}' 
                    AND G1_REVFIM <> 'ZZZ' AND struct.D_E_L_E_T_ <> '*' 
                    AND G1_REVFIM = (SELECT MAX(G1_REVFIM) FROM {database}.dbo.SG1010 WHERE G1_COD = '{codigo}'AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*')
                    ORDER BY B1_DESC ASC;
                """

                try:
                    conn_estrutura = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')

                    cursor_estrutura = conn_estrutura.cursor()
                    cursor_estrutura.execute(select_query_estrutura)
                    
                    nova_guia_estrutura = QWidget()
                    layout_nova_guia_estrutura = QVBoxLayout()
                    layout_cabecalho = QHBoxLayout()

                    tree_estrutura = QTableWidget(nova_guia_estrutura)
                    tree_estrutura.setColumnCount(len(cursor_estrutura.description))
                    tree_estrutura.setHorizontalHeaderLabels([desc[0] for desc in cursor_estrutura.description])
                    
                    # Tornar a tabela somente leitura
                    tree_estrutura.setEditTriggers(QTableWidget.NoEditTriggers)

                    # Permitir edição apenas na coluna "Quantidade" (assumindo que "Quantidade" é a terceira coluna, índice 2)
                    tree_estrutura.setEditTriggers(QAbstractItemView.DoubleClicked)
                    tree_estrutura.setItemDelegateForColumn(2, QItemDelegate(tree_estrutura))
                    
                    # Configurar a fonte da tabela
                    fonte_tabela = QFont("Segoe UI", 8)  # Substitua por sua fonte desejada e tamanho
                    tree_estrutura.setFont(fonte_tabela)

                    # Ajustar a altura das linhas
                    altura_linha = 22  # Substitua pelo valor desejado
                    tree_estrutura.verticalHeader().setDefaultSectionSize(altura_linha)

                    for i, row in enumerate(cursor_estrutura.fetchall()):
                        tree_estrutura.insertRow(i)
                        for j, value in enumerate(row):
                            if j == 2:
                                valor_formatado = "{:.2f}".format(float(value))
                            elif j == 5:
                                data_obj = datetime.strptime(value, "%Y%m%d")   
                                valor_formatado = data_obj.strftime("%d/%m/%Y")                  
                            else:
                                valor_formatado = str(value).strip()
                            
                            item = QTableWidgetItem(valor_formatado)
                            item.setForeground(QColor("#EEEEEE"))  # Definir cor do texto da coluna quantidade
                            
                            if j != 0 and j != 1:
                                item.setTextAlignment(Qt.AlignCenter)
                            
                            tree_estrutura.setItem(i, j, item)

                    tree_estrutura.setSortingEnabled(True)
                    
                    # Ajustar automaticamente a largura da coluna "Descrição"
                    self.ajustar_largura_coluna_descricao(tree_estrutura)
                        
                    layout_cabecalho.addWidget(QLabel(f"CONSULTA DE ESTRUTURA\n\n{codigo} - {descricao_onde_usado}"), alignment=Qt.AlignLeft)
                    layout_nova_guia_estrutura.addLayout(layout_cabecalho)                
                    layout_nova_guia_estrutura.addWidget(tree_estrutura)              
                    nova_guia_estrutura.setLayout(layout_nova_guia_estrutura)
                    
                    nova_guia_estrutura.setStyleSheet("""                                           
                        * {
                            background-color: #262626;
                        }
                        
                        QLabel {
                            color: #A7A6A6;
                            font-size: 18px;
                            font-weight: bold;
                        }
                        
                        QTableWidget {
                            border: 1px solid #000000;
                        }

                        QTableWidget QHeaderView::section {
                            background-color: #575a5f;
                            color: #fff;
                            padding: 5px;
                            height: 18px;
                        }

                        QTableWidget QHeaderView::section:horizontal {
                            border-top: 1px solid #333;
                        }
                        
                        QTableWidget::item:selected {
                            background-color: #0066ff;
                            color: #fff;
                            font-weight: bold;
                        }        
                    """)

                    if not self.existe_guias_abertas():
                        # Se não houver guias abertas, adicione a guia ao layout principal
                        self.layout().addWidget(self.tabWidget)
                        self.tabWidget.setVisible(True)
                        
                    self.tabWidget.addTab(nova_guia_estrutura, f"{codigo}")

                except pyodbc.Error as ex:
                    print(f"Falha na consulta de estrutura. Erro: {str(ex)}")

                finally:
                    self.tabWidget.setCurrentIndex(self.tabWidget.indexOf(nova_guia_estrutura))
                    tree_estrutura.itemChanged.connect(lambda item: self.handle_item_change(item, tree_estrutura, codigo))
                    self.guias_abertas.append(codigo)
                    conn_estrutura.close()

                    
    def alterar_quantidade_estrutura(self, codigo_pai, codigo_filho, quantidade):
        query_alterar_quantidade_estrutura = f"""UPDATE {database}.dbo.SG1010 SET G1_QUANT = {quantidade} WHERE G1_COD = '{codigo_pai}' AND G1_COMP = '{codigo_filho}'
                AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*'
                AND G1_REVFIM = (SELECT MAX(G1_REVFIM) FROM {database}.dbo.SG1010 WHERE G1_COD = '{codigo_pai}' AND G1_REVFIM <> 'ZZZ' AND D_E_L_E_T_ <> '*');
            """  
        try:
            with pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}') as conn:
                cursor = conn.cursor()
                cursor.execute(query_alterar_quantidade_estrutura)
                conn.commit()
            
        except Exception as ex:
            ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}", "Erro de execução", 16 | 0)
        
        
    def handle_item_change(self, item, tree_estrutura, codigo_pai):
        if item.column() == 2:    
            linha_selecionada = tree_estrutura.currentItem()
            
            codigo_filho = tree_estrutura.item(linha_selecionada.row(), 0).text()
            nova_quantidade = item.text()
            nova_quantidade = nova_quantidade.replace(',', '.')
            
            if nova_quantidade.replace('.', '', 1).isdigit():
                self.alterar_quantidade_estrutura(codigo_pai, codigo_filho, float(nova_quantidade))
            else:
                ctypes.windll.user32.MessageBoxW(
            0, "QUANTIDADE INVÁLIDA\n\nOs valores devem ser números, não nulos, sem espaços em branco e maiores que zero.\nPor favor, corrija tente novamente!", "SMARTPLIC®", 48 | 0)
    
    
    def executar_consulta_onde_usado(self):
        item_selecionado = self.tree.currentItem()

        if item_selecionado:
            codigo = self.tree.item(item_selecionado.row(), 0).text()
            descricao_onde_usado = self.tree.item(item_selecionado.row(), 1).text()
            
            if codigo not in self.guias_abertas_onde_usado:
                query_onde_usado = f"""
                    SELECT STRUT.G1_COD AS "CÓDIGO", PROD.B1_DESC "DESCRIÇÃO" 
                    FROM {database}.dbo.SG1010 STRUT 
                    INNER JOIN {database}.dbo.SB1010 PROD 
                    ON G1_COD = B1_COD WHERE G1_COMP = '{codigo}' 
                    AND STRUT.D_E_L_E_T_ <> '*';
                """
                self.guias_abertas_onde_usado.append(codigo)
                try:
                    conn_estrutura = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')

                    cursor_estrutura = conn_estrutura.cursor()
                    cursor_estrutura.execute(query_onde_usado)
                    
                    nova_guia_estrutura = QWidget()
                    layout_nova_guia_estrutura = QVBoxLayout()
                    layout_cabecalho = QHBoxLayout()

                    tree_estrutura = QTableWidget(nova_guia_estrutura)
                    tree_estrutura.setColumnCount(len(cursor_estrutura.description))
                    tree_estrutura.setHorizontalHeaderLabels([desc[0] for desc in cursor_estrutura.description])
                    
                    # Tornar a tabela somente leitura
                    tree_estrutura.setEditTriggers(QTableWidget.NoEditTriggers)
                    
                    # Configurar a fonte da tabela
                    fonte_tabela = QFont("Segoe UI", 8)  # Substitua por sua fonte desejada e tamanho
                    tree_estrutura.setFont(fonte_tabela)

                    # Ajustar a altura das linhas
                    altura_linha = 22  # Substitua pelo valor desejado
                    tree_estrutura.verticalHeader().setDefaultSectionSize(altura_linha)

                    for i, row in enumerate(cursor_estrutura.fetchall()):
                        tree_estrutura.insertRow(i)
                        for j, value in enumerate(row):
                            valor_formatado = str(value).strip()
                            
                            item = QTableWidgetItem(valor_formatado)                          
                            tree_estrutura.setItem(i, j, item)

                    tree_estrutura.setSortingEnabled(True)
                    
                    # Ajustar automaticamente a largura da coluna "Descrição"
                    self.ajustar_largura_coluna_descricao(tree_estrutura)
                        
                    layout_cabecalho.addWidget(QLabel(f'Onde é usado?\n\n{codigo} - {descricao_onde_usado}'), alignment=Qt.AlignLeft)
                    layout_nova_guia_estrutura.addLayout(layout_cabecalho)                
                    layout_nova_guia_estrutura.addWidget(tree_estrutura)              
                    nova_guia_estrutura.setLayout(layout_nova_guia_estrutura)
                    
                    nova_guia_estrutura.setStyleSheet("""                                           
                        * {
                            background-color: #262626;
                        }
                        
                        QLabel {
                            color: #A7A6A6;
                            font-size: 18px;
                            font-weight: bold;
                        }
                        
                        QTableWidget {
                            border: 1px solid #000000;
                        }

                        QTableWidget QHeaderView::section {
                            background-color: #575a5f;
                            color: #fff;
                            padding: 5px;
                            height: 18px;
                        }

                        QTableWidget QHeaderView::section:horizontal {
                            border-top: 1px solid #333;
                        }
                        
                        QTableWidget::item:selected {
                            background-color: #0066ff;
                            color: #fff;
                            font-weight: bold;
                        }        
                    """)

                    if not self.existe_guias_abertas():
                        # Se não houver guias abertas, adicione a guia ao layout principal
                        self.layout().addWidget(self.tabWidget)
                        self.tabWidget.setVisible(True)
                    
                    self.tabWidget.addTab(nova_guia_estrutura, f"Onde é usado? - {codigo}")
                    tree_estrutura.itemDoubleClicked.connect(self.copiar_linha)

                except pyodbc.Error as ex:
                    print(f"Falha na consulta de estrutura. Erro: {str(ex)}")

                finally:
                    self.tabWidget.setCurrentIndex(self.tabWidget.indexOf(nova_guia_estrutura))
                    conn_estrutura.close()


if __name__ == "__main__":
    # Parâmetros de conexão com o banco de dados SQL Server
    server = 'SVRERP,1433'
    database = 'PROTHEUS1233_HML' # PROTHEUS12_R27 (base de produção) PROTHEUS1233_HML (base de desenvolvimento/teste)
    username = 'coognicao'
    password = '0705@Abc'
    driver = '{ODBC Driver 17 for SQL Server}'

    app = QApplication(sys.argv)
    window = ConsultaApp()

    largura_janela = 1200  # Substitua pelo valor desejado
    altura_janela = 700 # Substitua pelo valor desejado

    largura_tela = app.primaryScreen().size().width()
    altura_tela = app.primaryScreen().size().height()

    pos_x = (largura_tela - largura_janela) // 2
    pos_y = (altura_tela - altura_janela) // 2

    window.setGeometry(pos_x, pos_y, largura_janela, altura_janela)

    window.show()
    sys.exit(app.exec_())
