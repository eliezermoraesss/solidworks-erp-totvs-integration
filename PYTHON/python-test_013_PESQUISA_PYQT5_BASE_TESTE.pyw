import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QHBoxLayout, QTableWidget, \
    QTableWidgetItem, QHeaderView, QSizePolicy, QSpacerItem, QMessageBox
from PyQt5.QtGui import QFont, QIcon, QDesktopServices, QColor
from PyQt5.QtCore import Qt, QUrl, QCoreApplication
import pyodbc
import pyperclip
import os
import time

class ConsultaApp(QWidget):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("CONSULTA DE PRODUTOS - TOTVS - BASE DE TESTES")
        
        # Configurar o ícone da janela
        icon_path = "010.png"
        self.setWindowIcon(QIcon(icon_path))
        
        # Ajuste a cor de fundo da janela
        self.setAutoFillBackground(True)
        palette = self.palette()
        palette.setColor(self.backgroundRole(), QColor('#eeeeee'))  # Substitua pela cor desejada
        self.setPalette(palette)
            
        # self.setWindowFlags(Qt.WindowStaysOnTopHint) # Exibir a janela sempre sobrepondo as demais janelas
        self.nova_janela = None  # Adicione esta linha
        
        # Aplicar folha de estilo ao aplicativo
        self.setStyleSheet("""                                                                         
            QLabel {
                color: #000;
                font-size: 11px;
                font-weight: bold;
            }

            QLineEdit {
                background-color: #fff;
                border: 1px solid #5ab3ee;
                padding: 5px;
                border: none;
                border-radius: 5px;
            }

            QPushButton {
                background-color: #0c9af8;
                color: #fff;
                padding: 5px 15px;
                border: 2px;
                border-radius: 5px;
                font-size: 11px;
                height: 20px;
                font-weight: bold;
                margin-top: 3px;
                margin-bottom: 3px;
            }

            QPushButton:hover {
                background-color: #2416e0;  /* Nova cor ao passar o mouse sobre o botão */
            }

            QPushButton:pressed {
                background-color: #ef08f7;  /* Nova cor ao clicar no botão */
            }

            QTableWidget {
                border: 1px solid #85aaf0;
            }

            QTableWidget QHeaderView::section {
                background-color: #7d85f0;
                color: #fff;
                padding: 5px;
            }

            QTableWidget QHeaderView::section:horizontal {
                border-top: 1px solid #333;
            }
            
            QTableWidget::item:selected {
                background-color: #8EC4FF; /* Altere para a cor desejada */
                color: #000; /* Cor do texto no item selecionado */
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

        self.configurar_campos()

        self.btn_consultar = QPushButton("Pesquisar", self)
        self.btn_consultar.clicked.connect(self.executar_consulta)
        self.btn_consultar.setMinimumWidth(100)  # Definindo o comprimento mínimo
        
        self.btn_limpar = QPushButton("Limpar", self)
        self.btn_limpar.clicked.connect(self.limpar_campos)
        self.btn_limpar.setMinimumWidth(100)  # Definindo o comprimento mínimo
        
        self.btn_nova_janela = QPushButton("Nova Janela", self)
        self.btn_nova_janela.clicked.connect(self.abrir_nova_janela)
        self.btn_nova_janela.setMinimumWidth(100)  # Definindo o comprimento mínimo
        
        self.btn_abrir_desenho = QPushButton("Abrir Desenho", self)
        self.btn_abrir_desenho.clicked.connect(self.abrir_desenho)
        self.btn_abrir_desenho.setMinimumWidth(100)  # Definindo o comprimento mínimo
        
        self.btn_fechar = QPushButton("Fechar", self)
        self.btn_fechar.clicked.connect(self.fechar_janela)
        self.btn_fechar.setMinimumWidth(100)  # Definindo o comprimento mínimo

        self.configurar_tabela()

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

        layout_linha_01.addWidget(QLabel("Descrição:"))
        layout_linha_01.addWidget(self.descricao_var)

        layout_linha_01.addWidget(QLabel("Contém na Descrição:"))
        layout_linha_01.addWidget(self.descricao2_var)

        layout_linha_02.addWidget(QLabel("Tipo:"))
        layout_linha_02.addWidget(self.tipo_var)

        layout_linha_02.addWidget(QLabel("Unidade de Medida:"))
        layout_linha_02.addWidget(self.um_var)

        layout_linha_02.addWidget(QLabel("Armazém:"))
        layout_linha_02.addWidget(self.armazem_var)

        layout_linha_02.addWidget(QLabel("Grupo:"))
        layout_linha_02.addWidget(self.grupo_var)
        
        layout_linha_02.addWidget(QLabel("Desc. Grupo:"))
        layout_linha_02.addWidget(self.grupo_desc_var)
        
        layout_linha_03.addWidget(self.btn_consultar)
        layout_linha_03.addWidget(self.btn_limpar)
        layout_linha_03.addWidget(self.btn_nova_janela)
        layout_linha_03.addWidget(self.btn_abrir_desenho)
        layout_linha_03.addWidget(self.btn_fechar)
        
        # Adicione um espaçador esticável para centralizar os botões
        layout_linha_03.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        
        layout.addLayout(layout_linha_01)
        layout.addLayout(layout_linha_02)
        layout.addLayout(layout_linha_03)
                
        layout.addWidget(self.tree)

        self.setLayout(layout)

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
        self.tree.setColumnCount(11)
        self.tree.setHorizontalHeaderLabels(
            ["CÓDIGO", "DESCRIÇÃO", "DESC. COMP.", "TIPO", "UM", "ARMAZÉM", "GRUPO", "DESC. GRUPO", "CC", "BLOQUEADO?",
             "REV."])
        self.tree.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tree.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tree.setSelectionBehavior(QTableWidget.SelectRows)
        self.tree.setSelectionMode(QTableWidget.SingleSelection)
        self.tree.setSortingEnabled(True)  # Permitir ordenação
        
        # Configurar a fonte da tabela
        fonte_tabela = QFont("Roboto", 8)  # Substitua por sua fonte desejada e tamanho
        self.tree.setFont(fonte_tabela)

        # Ajustar a altura das linhas
        altura_linha = 25  # Substitua pelo valor desejado
        self.tree.verticalHeader().setDefaultSectionSize(altura_linha)
        
        # Conectar o evento sectionClicked ao método ordenar_tabela
        self.tree.horizontalHeader().sectionClicked.connect(self.ordenar_tabela)
        
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

    def executar_consulta(self):
        # Obter os valores dos campos de consulta
        codigo = self.codigo_var.text().upper()
        descricao = self.descricao_var.text().upper()
        descricao2 = self.descricao2_var.text().upper()
        tipo = self.tipo_var.text().upper()
        um = self.um_var.text().upper()
        armazem = self.armazem_var.text().upper()
        grupo = self.grupo_var.text().upper()
        desc_grupo = self.grupo_desc_var.text().upper()

        # Construir a query de consulta
        select_query = f"""
        SELECT B1_COD, B1_DESC, B1_XDESC2, B1_TIPO, B1_UM, B1_LOCPAD, B1_GRUPO, B1_ZZNOGRP, B1_CC, B1_MSBLQL, B1_REVATU
        FROM PROTHEUS1233_HML.dbo.SB1010
        WHERE B1_COD LIKE '{codigo}%' AND B1_DESC LIKE '{descricao}%' AND B1_DESC LIKE '%{descricao2}%'
        AND B1_TIPO LIKE '{tipo}%' AND B1_UM LIKE '{um}%' AND B1_LOCPAD LIKE '{armazem}%' AND B1_GRUPO LIKE '{grupo}%' AND B1_ZZNOGRP LIKE '%{desc_grupo}%'
        """

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
            
            time.sleep(0.1)

            # Preencher a tabela com os resultados
            for i, row in enumerate(cursor.fetchall()):
                # Inserir os valores formatados na tabela
                self.tree.insertRow(i)
                for j, value in enumerate(row):
                    if j == 9:  # Verifica se o valor é da coluna B1_MSBLQL
                        # Converte o valor 1 para 'Sim' e 2 para 'Não'
                        value = 'Sim' if value == 1 else 'Não'
                    item = QTableWidgetItem(str(value).strip())
                    self.tree.setItem(i, j, item)
                    
                # Permitir que a interface gráfica seja atualizada
                QCoreApplication.processEvents()
                
            # Calcular a largura total das colunas
            largura_total_colunas = self.tree.horizontalHeader().length()

            # Definir a largura mínima da janela (opcional)
            largura_minima_janela = 800

            # Ajustar a largura da janela
            nova_largura_janela = max(largura_minima_janela, largura_total_colunas + 100)  # Adicione uma folga de 20 pixels
            self.setFixedWidth(nova_largura_janela)

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

    def copiar_linha(self):
        item_clicado = self.tree.currentItem()
        if item_clicado:
            valor_campo = item_clicado.text()
            pyperclip.copy(str(valor_campo))
            
    def abrir_nova_janela(self):
        if not self.nova_janela or not self.nova_janela.isVisible():
            self.nova_janela = ConsultaApp()
            self.nova_janela.setGeometry(self.x() + 50, self.y() + 50, self.width(), self.height())
            self.nova_janela.show()
            
    def fechar_janela(self):
        self.close()

if __name__ == "__main__":
    # Parâmetros de conexão com o banco de dados SQL Server
    server = 'SVRERP,1433'
    database = 'PROTHEUS1233_HML'
    username = 'coognicao'
    password = '0705@Abc'
    driver = '{ODBC Driver 17 for SQL Server}'

    app = QApplication(sys.argv)
    window = ConsultaApp()

    largura_janela = 1024  # Substitua pelo valor desejado
    altura_janela = 600 # Substitua pelo valor desejado

    largura_tela = app.primaryScreen().size().width()
    altura_tela = app.primaryScreen().size().height()

    pos_x = (largura_tela - largura_janela) // 2
    pos_y = (altura_tela - altura_janela) // 2

    window.setGeometry(pos_x, pos_y, largura_janela, altura_janela)

    window.show()
    sys.exit(app.exec_())