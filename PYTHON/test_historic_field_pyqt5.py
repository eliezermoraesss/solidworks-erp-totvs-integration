from PyQt5.QtWidgets import QLabel, QComboBox, QVBoxLayout, QWidget, QApplication
from PyQt5.QtCore import QSettings

class Janela(QWidget):
    def __init__(self):
        super().__init__()

        self.label = QLabel("Valor:")
        self.combo_historico = QComboBox()
        self.settings = QSettings("MinhaEmpresa", "MeuApp")
        self.nome_chave_historico = "historico_label_valor"

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.combo_historico)
        self.setLayout(layout)

        self.carregar_historico()

    def atualizar_valor(self, novo_valor):
        self.label.setText("Valor: " + novo_valor)

        historico = self.settings.value(self.nome_chave_historico, [])
        if novo_valor not in historico: # Evita duplicados
            historico.append(novo_valor)
            if len(historico) > 10: # Limita o histórico a 10 itens
                historico.pop(0)
            self.settings.setValue(self.nome_chave_historico, historico)
            self.combo_historico.clear()
            self.combo_historico.addItems(historico)

    def carregar_historico(self):
        historico = self.settings.value(self.nome_chave_historico, [])
        self.combo_historico.addItems(historico)

if __name__ == "__main__":
    app = QApplication([])
    janela = Janela()
    janela.show()
    janela.atualizar_valor("Valor inicial")
    janela.atualizar_valor("Segundo valor")
    janela.atualizar_valor("Terceiro valor")
    app.exec()
