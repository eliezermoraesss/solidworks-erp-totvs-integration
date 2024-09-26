import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QListWidget, QPushButton, QLineEdit, QTextEdit, QLabel, QHBoxLayout


class MultiSelectApp(QWidget):
    def __init__(self):
        super().__init__()

        # Configuração da janela
        self.setWindowTitle("Seleção Múltipla com Exibição e Ordenação")
        self.setGeometry(100, 100, 400, 500)

        # Layout principal
        layout = QVBoxLayout()

        # Variável para controlar a ordenação
        self.ordenacao_crescente = True

        # Botão para ordenar os itens
        btn_ordenar = QPushButton("Ordenar Itens (Crescente)")
        btn_ordenar.clicked.connect(self.ordenar_itens)
        layout.addWidget(btn_ordenar)

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

        # Campo de texto para exibir as seleções
        self.resultado_texto = QTextEdit()
        self.resultado_texto.setReadOnly(True)
        layout.addWidget(self.resultado_texto)

        # Botão para capturar seleção
        btn_capturar = QPushButton("Capturar Seleção")
        btn_capturar.clicked.connect(self.capturar_selecao)
        layout.addWidget(btn_capturar)

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

        # Exibir os itens selecionados no QTextEdit
        if itens_texto:
            self.resultado_texto.setText("\n".join(itens_texto))
        else:
            self.resultado_texto.setText("Nenhum item selecionado.")

    def ordenar_itens(self):
        # Ordenar os itens da lista
        if self.ordenacao_crescente:
            self.list_widget.sortItems()  # Ordena de forma crescente (padrão)
        else:
            # Obtem os itens, inverte a ordem e limpa a lista
            itens = [self.list_widget.item(i).text() for i in range(self.list_widget.count())]
            self.list_widget.clear()

            # Adiciona os itens em ordem decrescente
            self.list_widget.addItems(reversed(itens))

        # Alterna o estado da ordenação
        self.ordenacao_crescente = not self.ordenacao_crescente

        # Atualiza o texto do botão de ordenação
        btn_ordenar = self.sender()
        if self.ordenacao_crescente:
            btn_ordenar.setText("Ordenar Itens (Crescente)")
        else:
            btn_ordenar.setText("Ordenar Itens (Decrescente)")


# Inicializando a aplicação
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MultiSelectApp()
    window.show()
    sys.exit(app.exec_())
