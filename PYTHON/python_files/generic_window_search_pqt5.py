import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QLineEdit, QPushButton, QDialog, 
                             QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem, 
                             QHeaderView, QDialogButtonBox, QWidget)

class SearchDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Pesquisa")
        self.setMinimumWidth(400)

        # Campo de busca
        self.search_input = QLineEdit(self)
        self.search_input.setPlaceholderText("Digite o termo de pesquisa...")

        # Tabela de resultados
        self.table = QTableWidget(self)
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Código", "Descrição"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # Preenchendo a tabela com dados fictícios (você pode substituir por sua lógica de pesquisa)
        data = [
            ("001", "Item 1"),
            ("002", "Item 2"),
            ("003", "Item 3"),
        ]
        self.table.setRowCount(len(data))
        for row, (codigo, descricao) in enumerate(data):
            self.table.setItem(row, 0, QTableWidgetItem(codigo))
            self.table.setItem(row, 1, QTableWidgetItem(descricao))

        # Detecta clique duplo ou Enter na tabela
        self.table.itemDoubleClicked.connect(self.accept_selection)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        
        # Botões de Cancelar e Ok
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept_selection)
        buttons.rejected.connect(self.reject)

        # Layout da janela de diálogo
        layout = QVBoxLayout(self)
        layout.addWidget(self.search_input)
        layout.addWidget(self.table)
        layout.addWidget(buttons)

    def accept_selection(self):
        selected_items = self.table.selectedItems()
        if selected_items:
            codigo = selected_items[0].text()  # Pega o código da primeira coluna
            self.selected_code = codigo
            self.accept()

    def get_selected_code(self):
        return getattr(self, 'selected_code', None)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Exemplo com Lupa e Pesquisa")
        self.setMinimumWidth(400)

        # Criar um widget central explicitamente
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)  # Definindo o widget central

        # Campo A com botão de lupa
        self.input_a = QLineEdit(self)
        self.input_a.setPlaceholderText("Campo A")
        
        # Botão de lupa ao lado do campo
        self.search_button = QPushButton("🔍", self)
        self.search_button.setFixedWidth(40)
        self.search_button.clicked.connect(self.open_search_dialog)

        # Layout horizontal para campo A e o botão de lupa
        h_layout = QHBoxLayout()
        h_layout.addWidget(self.input_a)
        h_layout.addWidget(self.search_button)

        # Definir layout para o widget central
        layout = QVBoxLayout(central_widget)  # Criação do layout para o central_widget
        layout.addLayout(h_layout)

    def open_search_dialog(self):
        dialog = SearchDialog(self)
        if dialog.exec() == QDialog.Accepted:
            selected_code = dialog.get_selected_code()
            if selected_code:
                self.input_a.setText(selected_code)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
