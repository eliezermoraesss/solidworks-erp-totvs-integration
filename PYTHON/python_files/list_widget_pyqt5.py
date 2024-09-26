import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QListWidget, QPushButton, QLineEdit, QMessageBox, QHBoxLayout


class MultiSelectApp(QWidget):
    def __init__(self):
        super().__init__()

        # Configuração da janela
        self.setWindowTitle("Seleção Múltipla com Pesquisa e Ordenação")
        self.setGeometry(100, 100, 400, 500)

        # Layout principal
        layout = QVBoxLayout()

        # Campo de pesquisa
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Pesquisar...")
        self.search_box.textChanged.connect(self.filtrar_itens)  # Filtrar enquanto o usuário digita
        layout.addWidget(self.search_box)

        # Criação do QListWidget com seleção múltipla
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QListWidget.MultiSelection)

        # Adicionando muitos itens à lista (exemplo com 200 itens)
        self.itens_originais = [f"Item {i}" for i in range(1, 201)]
        self.list_widget.addItems(self.itens_originais)

        layout.addWidget(self.list_widget)

        # Botões para capturar seleção e ordenar
        btn_layout = QHBoxLayout()

        btn_capturar = QPushButton("Capturar Seleção")
        btn_capturar.clicked.connect(self.capturar_selecao)
        btn_layout.addWidget(btn_capturar)

        btn_ordenar = QPushButton("Ordenar Itens")
        btn_ordenar.clicked.connect(self.ordenar_itens)
        btn_layout.addWidget(btn_ordenar)

        layout.addLayout(btn_layout)

        # Definir o layout na janela
        self.setLayout(layout)

    def filtrar_itens(self):
        # Filtrar os itens com base no texto de pesquisa
        termo_pesquisa = self.search_box.text().lower()
        self.list_widget.clear()

        # Filtrar itens da lista original
        itens_filtrados = [item for item in self.itens_originais if termo_pesquisa in item.lower()]
        self.list_widget.addItems(itens_filtrados)

    def capturar_selecao(self):
        # Obter os itens selecionados
        itens_selecionados = self.list_widget.selectedItems()

        # Pegar os textos dos itens selecionados
        itens_texto = [item.text() for item in itens_selecionados]

        # Exibir os itens selecionados em uma caixa de mensagem
        if itens_texto:
            QMessageBox.information(self, "Itens Selecionados", ", ".join(itens_texto))
        else:
            QMessageBox.warning(self, "Nenhum item selecionado", "Por favor, selecione pelo menos um item.")

    def ordenar_itens(self):
        # Ordenar os itens da lista
        self.list_widget.sortItems()

# Inicializando a aplicação
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MultiSelectApp()
    window.show()
    sys.exit(app.exec_())
