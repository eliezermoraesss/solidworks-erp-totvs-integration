
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QHBoxLayout, QTableWidget, \
    QHeaderView, QSizePolicy, QSpacerItem, QTableView
from PyQt5.QtGui import QFont, QIcon, QDesktopServices, QStandardItemModel
from PyQt5.QtCore import Qt, QUrl, QCoreApplication
import pyodbc
import pyperclip
import os

class ResultadosTableModel(QStandardItemModel):
    def __init__(self, data, header_labels):
        super().__init__()
        self._data = data
        self._header_labels = header_labels
        self.setHorizontalHeaderLabels(header_labels)
        self.carregar_dados_iniciais()

    def rowCount(self, parent):
        return len(self._data)

    def columnCount(self, parent):
        return len(self._header_labels)

    def carregar_dados_iniciais(self):
        # Adicione aqui a lógica para carregar os dados iniciais
        self.inserir_dados_iniciais()

    def inserir_dados_iniciais(self):
        # Adicione aqui a lógica para inserir os dados iniciais na tabela
        pass

class ConsultaApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Consulta de Produtos TOTVS")
        self.setWindowIcon(QIcon(r'//192.175.175.4/f/ELIEZER/PROJETO SOLIDWORKS TOTVS/VBA/assets/images/004.png'))

        self.nova_janela = None

        self.setStyleSheet("""
            QLabel {
                color: #333;
            }

            QLineEdit {
                background-color: #fff;
                border: 1px solid #ccc;
                padding: 5px;
                border: none;
                border-radius: 3px;
            }

            QPushButton {
                background-color: #0c9af8;
                color: #fff;
                padding: 5px 10px;
                border: none;
                border-radius: 3px;
            }

            QPushButton:hover {
                background-color: #2416e0;
            }

            QPushButton:pressed {
                background-color: #ef08f7;
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
                border-top: 1px solid #ccc;
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
        self.btn_consultar.setMinimumWidth(100)

        self.btn_limpar = QPushButton("Limpar", self)
        self.btn_limpar.clicked.connect(self.limpar_campos)
        self.btn_limpar.setMinimumWidth(100)

        self.btn_nova_janela = QPushButton("Nova Janela", self)
        self.btn_nova_janela.clicked.connect(self.abrir_nova_janela)
        self.btn_nova_janela.setMinimumWidth(100)

        self.btn_fechar = QPushButton("Fechar", self)
        self.btn_fechar.clicked.connect(self.fechar_janela)
        self.btn_fechar.setMinimumWidth(100)

        self.configurar_tabela()

        self.offset = 0
        self.limit = 500

        self.model = ResultadosTableModel([], ["CÓDIGO", "DESCRIÇÃO", "DESC. COMP.", "TIPO", "UM", "ARMAZÉM", "GRUPO", "DESC. GRUPO", "CC", "BLOQUEADO?", "REV."])
        self.tree.setModel(self.model)

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
        layout_linha_03.addWidget(self.btn_fechar)

        layout_linha_03.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))

        layout.addLayout(layout_linha_01)
        layout.addLayout(layout_linha_02)
        layout.addLayout(layout_linha_03)

        layout.addWidget(self.tree)

        self.setLayout(layout)

    def configurar_campos(self):
        self.codigo_var.setMinimumWidth(80)
        self.descricao_var.setMinimumWidth(200)
        self.descricao2_var.setMinimumWidth(200)
        self.tipo_var.setMinimumWidth(200)
        self.um_var.setMinimumWidth(200)
        self.armazem_var.setMinimumWidth(200)
        self.grupo_var.setMinimumWidth(200)
        self.grupo_desc_var.setMinimumWidth(200)

    def configurar_tabela(self):
        self.tree = QTableView(self)
        self.tree.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tree.setEditTriggers(QTableView.NoEditTriggers)
        self.tree.setSelectionBehavior(QTableView.SelectRows)
        self.tree.setSelectionMode(QTableView.SingleSelection)
        self.tree.setSortingEnabled(True)

        fonte_tabela = QFont("Arial", 8)
        self.tree.setFont(fonte_tabela)

        altura_linha = 25
        self.tree.verticalHeader().setDefaultSectionSize(altura_linha)

        self.tree.horizontalHeader().sectionClicked.connect(self.ordenar_tabela)
        self.tree.clicked.connect(self.tree_item_clicked)
        self.tree.verticalScrollBar().valueChanged.connect(self.carregar_mais_dados)

    def ordenar_tabela(self, logicalIndex):
        index = self.tree.horizontalHeader().sortIndicatorOrder()
        order = Qt.AscendingOrder if index == 0 else Qt.DescendingOrder
        self.tree.sortByColumn(logicalIndex, order)

    def limpar_campos(self):
        self.codigo_var.clear()
        self.descricao_var.clear()
        self.descricao2_var.clear()
        self.tipo_var.clear()
        self.um_var.clear()
        self.armazem_var.clear()
        self.grupo_var.clear()
        self.grupo_desc_var.clear()

    def executar_consulta(self):
        codigo = self.codigo_var.text().upper()
        descricao = self.descricao_var.text().upper()
        descricao2 = self.descricao2_var.text().upper()
        tipo = self.tipo_var.text().upper()
        um = self.um_var.text().upper()
        armazem = self.armazem_var.text().upper()
        grupo = self.grupo_var.text().upper()
        desc_grupo = self.grupo_desc_var.text().upper()

        select_query = f"""
        SELECT B1_COD, B1_DESC, B1_XDESC2, B1_TIPO, B1_UM, B1_LOCPAD, B1_GRUPO, B1_ZZNOGRP, B1_CC, B1_MSBLQL, B1_REVATU
        FROM PROTHEUS12_R27.dbo.SB1010
        WHERE B1_COD LIKE '{codigo}%' AND B1_DESC LIKE '{descricao}%' AND B1_DESC LIKE '%{descricao2}%'
        AND B1_TIPO LIKE '{tipo}%' AND B1_UM LIKE '{um}%' AND B1_LOCPAD LIKE '{armazem}%' AND B1_GRUPO LIKE '{grupo}%' AND B1_ZZNOGRP LIKE '%{desc_grupo}%'
        ORDER BY B1_COD
        OFFSET {self.offset} ROWS
        FETCH NEXT {self.limit} ROWS ONLY
        """

        try:
            conn = pyodbc.connect(
                f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
            cursor = conn.cursor()
            cursor.execute(select_query)

            rows_to_insert = []
            for row in cursor.fetchall():
                rows_to_insert.append(row)

            # Atualizar o modelo com os novos dados
            current_rows = self.model._data
            self.model.beginInsertRows(self.model.index(len(current_rows), 0),
                                       len(current_rows), len(current_rows) + len(rows_to_insert) - 1)
            self.model._data.extend(rows_to_insert)
            self.model.endInsertRows()

            QCoreApplication.processEvents()

        except pyodbc.Error as ex:
            print(f"Falha na consulta. Erro: {str(ex)}")

        finally:
            conn.close()

    def tree_item_clicked(self, item):
        if item.column() == 10:
            codigo = self.tree.model().index(item.row(), 0).data()
            pdf_path = os.path.join("Y:/PDF-OFICIAL/", f"{codigo}.PDF")
            pdf_path = os.path.normpath(pdf_path)
            if os.path.exists(pdf_path):
                QCoreApplication.processEvents()
                QDesktopServices.openUrl(QUrl.fromLocalFile(pdf_path))
            else:
                print(f"Arquivo PDF não encontrado para o código {codigo}.")

    def carregar_mais_dados(self):
        scroll_bar = self.tree.verticalScrollBar()
        if scroll_bar.value() == scroll_bar.maximum():
            self.offset += self.limit
            self.executar_consulta()

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
    server = 'SVRERP,1433'
    database = 'PROTHEUS12_R27'
    username = 'coognicao'
    password = '0705@Abc'
    driver = '{ODBC Driver 17 for SQL Server}'

    app = QApplication(sys.argv)
    window = ConsultaApp()

    largura_janela = 1024
    altura_janela = 600

    largura_tela = app.primaryScreen().size().width()
    altura_tela = app.primaryScreen().size().height()

    pos_x = (largura_tela - largura_janela) // 2
    pos_y = (altura_tela - altura_janela) // 2

    window.setGeometry(pos_x, pos_y, largura_janela, altura_janela)

    window.show()
    sys.exit(app.exec_())
