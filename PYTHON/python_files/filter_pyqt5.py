import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QTableWidget, QTableWidgetItem, QLineEdit, QHeaderView
from PyQt5.QtCore import QSortFilterProxyModel, Qt

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Criação do layout principal
        self.setWindowTitle('Tabela com Filtro')
        self.setGeometry(300, 300, 600, 400)

        # Criação de um widget central
        widget = QWidget(self)
        self.setCentralWidget(widget)
        layout = QVBoxLayout(widget)

        # Campo de texto para o filtro
        self.filter_input = QLineEdit(self)
        self.filter_input.setPlaceholderText('Digite o valor para filtrar...')
        layout.addWidget(self.filter_input)

        # Criação do QTableWidget
        self.table_widget = QTableWidget(self)
        self.table_widget.setRowCount(10)
        self.table_widget.setColumnCount(3)
        self.table_widget.setHorizontalHeaderLabels(['Coluna 1', 'Coluna 2', 'Coluna 3'])

        # Configurar cabeçalho para redimensionamento automático
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # Inserção de dados na tabela
        data = [
            ['John', 'Doe', '30'],
            ['Jane', 'Doe', '25'],
            ['Mike', 'Brown', '45'],
            ['Anna', 'Smith', '32'],
            ['Joe', 'Davis', '50'],
            ['Chris', 'Johnson', '40'],
            ['Amanda', 'Clark', '35'],
            ['Nancy', 'Jones', '29'],
            ['Sarah', 'Miller', '22'],
            ['Tom', 'Wilson', '33']
        ]

        for row_idx, row_data in enumerate(data):
            for col_idx, value in enumerate(row_data):
                self.table_widget.setItem(row_idx, col_idx, QTableWidgetItem(value))

        layout.addWidget(self.table_widget)

        # Criação de um modelo de filtro
        self.proxy_model = QSortFilterProxyModel(self)
        self.proxy_model.setSourceModel(self.table_widget.model())
        self.proxy_model.setFilterCaseSensitivity(Qt.CaseInsensitive)

        # Conectar o campo de texto ao método de filtro
        self.filter_input.textChanged.connect(self.apply_filter)

    def apply_filter(self, text):
        """Aplica o filtro à tabela"""
        self.proxy_model.setFilterKeyColumn(-1)  # -1 para filtrar em todas as colunas
        self.proxy_model.setFilterFixedString(text)
        self.table_widget.setModel(self.proxy_model)

# Executa a aplicação
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
