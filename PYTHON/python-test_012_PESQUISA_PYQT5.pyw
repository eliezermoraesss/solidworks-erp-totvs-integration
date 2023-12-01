import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QTableWidget, QTableWidgetItem, QHeaderView
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtCore import Qt
import pyodbc
import pyperclip

class ConsultaApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Consulta de Produtos no TOTVS")
        self.setWindowIcon(QIcon(r'\\192.175.175.4\f\ELIEZER\PROJETO SOLIDWORKS TOTVS\VBA\assets\images\lupa.ico'))

        self.codigo_var = QLineEdit(self)
        self.descricao_var = QLineEdit(self)
        self.descricao2_var = QLineEdit(self)

        self.configurar_campos()

        self.btn_consultar = QPushButton("Pesquisar", self)
        self.btn_consultar.clicked.connect(self.executar_consulta)

        self.configurar_tabela()
        
        # Conectar o evento returnPressed dos campos de entrada ao método executar_consulta
        self.codigo_var.returnPressed.connect(self.executar_consulta)
        self.descricao_var.returnPressed.connect(self.executar_consulta)
        self.descricao2_var.returnPressed.connect(self.executar_consulta)

        layout = QVBoxLayout()
        layout.addWidget(QLabel("Código:"))
        layout.addWidget(self.codigo_var)
        layout.addWidget(QLabel("Descrição:"))
        layout.addWidget(self.descricao_var)
        layout.addWidget(QLabel("Contém na Descrição:"))
        layout.addWidget(self.descricao2_var)
        layout.addWidget(self.btn_consultar)
        layout.addWidget(self.tree)

        self.setLayout(layout)

    def configurar_campos(self):
        self.codigo_var.setMinimumWidth(200)
        self.descricao_var.setMinimumWidth(200)
        self.descricao2_var.setMinimumWidth(200)

    def configurar_tabela(self):
        self.tree = QTableWidget(self)
        self.tree.setColumnCount(11)
        self.tree.setHorizontalHeaderLabels(["CÓDIGO", "DESCRIÇÃO", "DESC. COMP.", "TIPO", "UM", "ARMAZÉM", "GRUPO", "DESC. GRUPO", "CC", "BLOQUEADO?", "REV."])
        self.tree.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tree.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tree.setSelectionBehavior(QTableWidget.SelectRows)
        self.tree.setSelectionMode(QTableWidget.SingleSelection)

    def executar_consulta(self):
        # Obter os valores dos campos de consulta
        codigo = self.codigo_var.text().upper()
        descricao = self.descricao_var.text().upper()
        descricao2 = self.descricao2_var.text().upper()

        # Construir a query de consulta
        select_query = f"""
        SELECT B1_COD, B1_DESC, B1_XDESC2, B1_TIPO, B1_UM, B1_LOCPAD, B1_GRUPO, B1_ZZNOGRP, B1_CC, B1_MSBLQL, B1_REVATU
        FROM PROTHEUS12_R27.dbo.SB1010
        WHERE B1_COD LIKE '{codigo}%' AND B1_DESC LIKE '{descricao}%' AND B1_DESC LIKE '%{descricao2}%'
        """

        try:
            # Estabelecer a conexão com o banco de dados
            conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')

            # Criar um cursor para executar comandos SQL
            cursor = conn.cursor()

            # Executar a consulta
            cursor.execute(select_query)

            # Limpar a tabela
            self.tree.setRowCount(0)

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

        except pyodbc.Error as ex:
            print(f"Falha na consulta. Erro: {str(ex)}")

        finally:
            # Fechar a conexão com o banco de dados
            conn.close()

    def copiar_linha(self):
        item_clicado = self.tree.currentItem()
        if item_clicado:
            valor_campo = item_clicado.text()
            pyperclip.copy(str(valor_campo))


if __name__ == "__main__":
    # Parâmetros de conexão com o banco de dados SQL Server
    server = 'SVRERP,1433'
    database = 'PROTHEUS12_R27'
    username = 'coognicao'
    password = '0705@Abc'
    driver = '{ODBC Driver 17 for SQL Server}'

    app = QApplication(sys.argv)
    window = ConsultaApp()
    
    largura_janela = 1204  # Substitua pelo valor desejado
    altura_janela = 768  # Substitua pelo valor desejado

    largura_tela = app.primaryScreen().size().width()
    altura_tela = app.primaryScreen().size().height()

    pos_x = (largura_tela - largura_janela) // 2
    pos_y = (altura_tela - altura_janela) // 2

    window.setGeometry(pos_x, pos_y, largura_janela, altura_janela)
    
    window.show()
    sys.exit(app.exec_())
