import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QPushButton, QVBoxLayout, QWidget


class CascadingTable(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Tabela em Cascata")
        self.setGeometry(100, 100, 800, 600)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.table = QTableWidget(self)
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Produto", "Detalhes"])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)  # Impede a edição da tabela

        self.layout = QVBoxLayout(self.central_widget)
        self.layout.addWidget(self.table)

        self.populate_table()

    def populate_table(self):
        # Exemplo de dados, você deve substituir isso pelos seus próprios dados
        data = [
            {"Produto": "Produto 1", "Detalhes": "Detalhes do Produto 1"},
            {"Produto": "Produto 2", "Detalhes": "Detalhes do Produto 2"},
            {"Produto": "Produto 3", "Detalhes": "Detalhes do Produto 3"},
            {"Produto": "Produto 4", "Detalhes": "Detalhes do Produto 4"},
            # Adicione mais dados conforme necessário
        ]

        for row, item in enumerate(data):
            self.add_table_row(row, item)

    def add_table_row(self, row, item):
        self.table.insertRow(row)

        product_item = QTableWidgetItem(item["Produto"])
        details_item = QTableWidgetItem(item["Detalhes"])

        self.table.setItem(row, 0, product_item)
        self.table.setItem(row, 1, details_item)

        # Adiciona botões de expandir/recolher
        expand_button = QPushButton("+", self)
        expand_button.clicked.connect(lambda _, r=row: self.toggle_row(r))
        self.table.setCellWidget(row, 2, expand_button)

    def toggle_row(self, row):
        current_text = self.table.cellWidget(row, 2).text()

        if current_text == "+":
            self.table.cellWidget(row, 2).setText("-")
            # Adicione aqui a lógica para expandir a linha, se necessário
        else:
            self.table.cellWidget(row, 2).setText("+")
            # Adicione aqui a lógica para recolher a linha, se necessário


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CascadingTable()
    window.show()
    sys.exit(app.exec_())
