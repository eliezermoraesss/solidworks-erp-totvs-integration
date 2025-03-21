import csv
import io
import locale
import os

import pandas as pd
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QIcon
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QTableWidget, QTableWidgetItem, QSplitter, QLineEdit, QLabel, QHBoxLayout,
                             QAbstractItemView, QAction, QMenu, QStyle, QComboBox)
from sqlalchemy import create_engine

from .db_mssql import setup_mssql
from .utils import abrir_desenho


def format_quantity(value):
    """Formata a quantidade: inteiro sem casas decimais, decimal com duas casas"""
    if value.is_integer():
        return f"{int(value)}"
    return locale.format_string("%.2f", value, grouping=True)


class BOMViewer(QMainWindow):
    def __init__(self, codigo_pai):
        super().__init__()
        self.engine = None
        self.codigo_pai = codigo_pai
        self.descricao_pai = None
        self.current_theme = "dark"  # Tema padrão
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
        self.all_components = []
        self.setWindowTitle(f'Eureka® - Visualizador de hierarquia de estrutura - {codigo_pai}')

        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Criar layout para filtro
        filter_layout = QHBoxLayout()

        # Filtro por Nível
        filter_label_nivel = QLabel("Nível:")
        self.filter_combobox_nivel = QComboBox()
        self.filter_combobox_nivel.addItem("")  # Nível padrão vazio
        self.filter_combobox_nivel.currentIndexChanged.connect(self.filter_tables)
        filter_layout.addWidget(filter_label_nivel)
        filter_layout.addWidget(self.filter_combobox_nivel)

        # Filtro por Código
        filter_label_codigo = QLabel("Código:")
        self.filter_input_codigo = QLineEdit()
        self.filter_input_codigo.setMaxLength(13)
        self.filter_input_codigo.setFixedWidth(150)  # Ajusta a largura conforme o número máximo de caracteres
        self.filter_input_codigo.textChanged.connect(self.filter_tables)
        self.add_clear_button(self.filter_input_codigo)  # Adiciona o botão de limpar com ícone
        filter_layout.addWidget(filter_label_codigo)
        filter_layout.addWidget(self.filter_input_codigo)

        # Filtro por Descrição (início)
        filter_label_desc = QLabel("Descrição:")
        self.filter_input_desc = QLineEdit()
        self.filter_input_desc.setMaxLength(100)
        self.filter_input_desc.textChanged.connect(self.filter_tables)
        self.add_clear_button(self.filter_input_desc)  # Adiciona o botão de limpar com ícone
        filter_layout.addWidget(filter_label_desc)
        filter_layout.addWidget(self.filter_input_desc)

        # Filtro por Descrição (contém)
        filter_label_desc_contains = QLabel("Contém na Descrição:")
        self.filter_input_desc_contains = QLineEdit()
        self.filter_input_desc_contains.setMaxLength(60)
        self.filter_input_desc_contains.textChanged.connect(self.filter_tables)
        self.add_clear_button(self.filter_input_desc_contains)  # Adiciona o botão de limpar com ícone
        filter_layout.addWidget(filter_label_desc_contains)
        filter_layout.addWidget(self.filter_input_desc_contains)

        layout.addLayout(filter_layout)

        # Splitter para dividir tabela e árvore
        splitter = QSplitter(Qt.Horizontal)

        # Tabela
        self.table = QTableWidget()
        self.table.setSortingEnabled(False)
        self.table.setAlternatingRowColors(True)
        splitter.addWidget(self.table)

        # Setup banco de dados e carregar dados
        self.setup_database()
        self.load_data()

    def setup_database(self):
        username, password, database, server = setup_mssql()
        driver = '{SQL Server}'
        conn_str = f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        self.engine = create_engine(f'mssql+pyodbc:///?odbc_connect={conn_str}')

    def get_components(self, parent_code: str, parent_qty: float = 1.0, level: int = 1):
        query = f"""
        SELECT 
            struct.G1_COD AS 'Código Pai',
            struct.G1_COMP AS 'Código',
            prod.B1_DESC AS 'Descrição',
            struct.G1_XUM AS 'Unidade',
            struct.G1_QUANT AS 'Quantidade'
        FROM 
            PROTHEUS12_R27.dbo.SG1010 struct
        INNER JOIN
            PROTHEUS12_R27.dbo.SB1010 prod
        ON 
            G1_COMP = B1_COD AND prod.D_E_L_E_T_ <> '*'
        WHERE 
            G1_COD = '{parent_code}' 
            AND G1_REVFIM <> 'ZZZ' 
            AND struct.D_E_L_E_T_ <> '*' 
            AND G1_REVFIM = (
                SELECT MAX(G1_REVFIM) 
                FROM PROTHEUS12_R27.dbo.SG1010 
                WHERE G1_COD = '{parent_code}'
                AND G1_REVFIM <> 'ZZZ' 
                AND struct.D_E_L_E_T_ <> '*'
            )
        """

        try:
            df = pd.read_sql(query, self.engine)

            if not df.empty:
                for _, row in df.iterrows():
                    total_qty = parent_qty * row['Quantidade']

                    component = {
                        'Nível': level,
                        'Código': row['Código'].strip(),
                        'Código Pai': row['Código Pai'].strip(),
                        'Descrição': row['Descrição'].strip(),
                        'Unidade': row['Unidade'].strip(),
                        'Quantidade': row['Quantidade'],
                        'Quantidade Total': total_qty
                    }
                    self.all_components.append(component)

                    self.get_components(row['Código'], total_qty, level + 1)

        except Exception as e:
            print(f"Erro ao processar código {parent_code}: {str(e)}")

    def load_data(self):
        codigo_pai = self.codigo_pai
        self.all_components = []

        # Adicionar item de nível mais alto
        # Primeiro precisamos buscar a descrição do item pai
        query = f"""
        SELECT B1_DESC as 'Descrição'
        FROM PROTHEUS12_R27.dbo.SB1010
        WHERE B1_COD = '{codigo_pai}'
        AND D_E_L_E_T_ <> '*'
        """
        try:
            df_pai = pd.read_sql(query, self.engine)
            descricao_pai = df_pai['Descrição'].iloc[0] if not df_pai.empty else ''
            self.descricao_pai = descricao_pai
        except:
            descricao_pai = ''

        self.all_components.append({
            'Nível': 0,
            'Código': codigo_pai.strip(),
            'Descrição': descricao_pai.strip(),
            'Código Pai': '-',
            'Unidade': 'UN',
            'Quantidade': 1.0,
            'Quantidade Total': 1.0
        })

        self.get_components(codigo_pai)
        self.df = pd.DataFrame(self.all_components)
        self.df.insert(5, 'Desenho PDF', '')
        # Reordenar as colunas
        self.df = self.df[['Nível', 'Código', 'Código Pai', 'Descrição', 'Desenho PDF', 'Unidade', 'Quantidade', 'Quantidade Total']]

        self.populate_nivel_combobox()
        self.setup_table()

    def populate_nivel_combobox(self):
        niveis = sorted(self.df['Nível'].unique())
        for nivel in niveis:
            self.filter_combobox_nivel.addItem(str(nivel))

    def setup_table(self):
        self.table.setColumnCount(len(self.df.columns))
        self.table.setHorizontalHeaderLabels(self.df.columns)
        self.populate_table(self.df)
        self.table.resizeColumnsToContents()

        # Tornar a tabela não editável
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        # Permitir seleção de múltiplas linhas
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)

        # Habilitar menu de contexto
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)

        # Conectar o sinal de duplo clique
        self.table.cellDoubleClicked.connect(self.copy_cell)

    def show_context_menu(self, position):
        menu = QMenu()
        copy_action = QAction("Copiar", self)
        copy_action.triggered.connect(self.copy_selection)

        open_drawing_action = QAction("Abrir desenho...", self)
        open_drawing_action.triggered.connect(lambda: abrir_desenho(self, self.table))

        menu.addAction(copy_action)
        menu.addAction(open_drawing_action)
        menu.exec_(self.table.viewport().mapToGlobal(position))

    def copy_selection(self):
        selection = self.table.selectedIndexes()
        if selection:
            rows = sorted(index.row() for index in selection)
            columns = sorted(index.column() for index in selection)
            rowcount = rows[-1] - rows[0] + 1
            colcount = columns[-1] - columns[0] + 1
            table = [[''] * colcount for _ in range(rowcount)]
            for index in selection:
                row = index.row() - rows[0]
                col = index.column() - columns[0]
                table[row][col] = self.table.item(index.row(), index.column()).text()
            stream = io.StringIO()
            csv.writer(stream, delimiter='\t').writerows(table)
            QApplication.clipboard().setText(stream.getvalue())

    def populate_table(self, df):
        self.table.setRowCount(len(df))
        COLOR_FILE_EXISTS = QColor(51, 211, 145)  # green
        COLOR_FILE_MISSING = QColor(201, 92, 118)  # light red
        for i, row in df.iterrows():
            for j, (column_name, value) in enumerate(row.items()):
                item = QTableWidgetItem()

                # Formatar quantidades (colunas Quantidade e Quantidade Total)
                if column_name in ['Quantidade', 'Quantidade Total']:
                    formatted_value = format_quantity(value)
                    item.setText(formatted_value)
                # Special handling for Desenho PDF column
                elif column_name == 'Desenho PDF':
                    codigo_desenho = row['Código'].strip()
                    pdf_path = os.path.join(r"\\192.175.175.4\dados\EMPRESA\PROJETOS\PDF-OFICIAL",
                                            f"{codigo_desenho}.PDF")
                    if os.path.exists(pdf_path):
                        item.setBackground(COLOR_FILE_EXISTS)
                        item.setText('Sim')
                        item.setToolTip("Desenho encontrado")
                    else:
                        item.setBackground(COLOR_FILE_MISSING)
                        item.setText('Não')
                        item.setToolTip("Desenho não encontrado")
                else:
                    item.setText(str(value))

                # Alinhar ao centro exceto código e descrição
                if column_name not in ['Código', 'Descrição']:
                    item.setTextAlignment(Qt.AlignCenter)
                else:
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)

                self.table.setItem(i, j, item)

    def filter_tables(self):
        filter_nivel = self.filter_combobox_nivel.currentText().strip()
        filter_codigo = self.filter_input_codigo.text().lower().strip()
        filter_desc = self.filter_input_desc.text().lower().strip()
        filter_desc_contains = self.filter_input_desc_contains.text().lower().strip()

        # PARTE 1: FILTRAGEM DA TABELA
        for row in range(self.table.rowCount()):
            row_visible = True
            nivel = self.table.item(row, 0).text().strip()
            codigo = self.table.item(row, 1).text().lower()
            descricao = self.table.item(row, 3).text().lower()

            if filter_nivel and filter_nivel != nivel:
                row_visible = False
            if filter_codigo and filter_codigo not in codigo:
                row_visible = False
            if filter_desc and not descricao.startswith(filter_desc):
                row_visible = False
            if filter_desc_contains and filter_desc_contains not in descricao:
                row_visible = False

            self.table.setRowHidden(row, not row_visible)

    def add_clear_button(self, line_edit):
        clear_icon = self.style().standardIcon(QStyle.SP_LineEditClearButton)

        line_edit_height = line_edit.height()
        pixmap = clear_icon.pixmap(line_edit_height - 4, line_edit_height - 4)
        larger_clear_icon = QIcon(pixmap)

        clear_action = QAction(larger_clear_icon, "Limpar", line_edit)
        clear_action.triggered.connect(line_edit.clear)
        line_edit.addAction(clear_action, QLineEdit.TrailingPosition)

    def copy_tree_item(self, item):
        if item:
            QApplication.clipboard().setText(item.text(0))

    def abrir_desenho_arvore(self, codigo):
        pdf_path = os.path.join(r"\\192.175.175.4\dados\EMPRESA\PROJETOS\PDF-OFICIAL",
                                f"{codigo}.PDF")

        if os.path.exists(pdf_path):
            os.startfile(pdf_path)

    def copy_cell(self, row, column):
        """Copia o conteúdo da célula selecionada para a área de transferência quando houver duplo clique"""
        item = self.table.item(row, column)
        if item:
            QApplication.clipboard().setText(item.text())
