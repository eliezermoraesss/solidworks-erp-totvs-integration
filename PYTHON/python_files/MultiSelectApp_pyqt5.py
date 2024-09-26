import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QListWidget, QPushButton, QMessageBox


class MultiSelectApp(QWidget):
    def __init__(self):
        super().__init__()

        # Configuração da janela
        self.setWindowTitle("Seleção Múltipla de Itens")
        self.setGeometry(100, 100, 300, 400)

        # Layout
        layout = QVBoxLayout()

        # Criação do QListWidget com seleção múltipla
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QListWidget.MultiSelection)

        # Adicionando itens à lista
        items = ["Item 1", "Item 2", "Item 3", "Item 4", "Item 5"]
        self.list_widget.addItems(items)

        # Botão para capturar seleções
        btn_capturar = QPushButton("Capturar Seleção")
        btn_capturar.clicked.connect(self.capturar_selecao)

        # Adicionar widgets ao layout
        layout.addWidget(self.list_widget)
        layout.addWidget(btn_capturar)

        # Definir o layout na janela
        self.setLayout(layout)

    def capturar_selecao(self):
        # Obter os itens selecionados
        itens_selecionados = self.list_widget.selectedItems()

        # Pegar os textos dos itens selecionados
        itens_texto = [item.text() for item in itens_selecionados]

        # Exibir em uma caixa de mensagem
        QMessageBox.information(self, "Itens Selecionados", ", ".join(itens_texto))

# Inicializando a aplicação
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MultiSelectApp()
    window.show()
    sys.exit(app.exec_())
