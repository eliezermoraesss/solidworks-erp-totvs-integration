import sys
from PyQt5.QtWidgets import QApplication, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget
from PyQt5.QtCore import Qt

class TabelaColorida(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Tabela com Linhas Alternadas de Cores')
        self.setGeometry(100, 100, 600, 400)

        self.layout = QVBoxLayout()

        # Criar uma tabela com 5 linhas e 3 colunas
        self.tabela = QTableWidget(self)
        self.tabela.setRowCount(5)
        self.tabela.setColumnCount(3)

        # Preencher a tabela com alguns dados de exemplo
        for row in range(5):
            for col in range(3):
                item = QTableWidgetItem(f'Dado {row}-{col}')
                self.tabela.setItem(row, col, item)

        # Definir cores alternadas para as linhas da tabela
        for row in range(self.tabela.rowCount()):
            if row % 2 == 0:
                # Linha par
                self.tabela.item(row, 0).setBackground(Qt.lightGray)
                self.tabela.item(row, 1).setBackground(Qt.lightGray)
                self.tabela.item(row, 2).setBackground(Qt.lightGray)
            else:
                # Linha ímpar
                self.tabela.item(row, 0).setBackground(Qt.white)
                self.tabela.item(row, 1).setBackground(Qt.white)
                self.tabela.item(row, 2).setBackground(Qt.white)

        self.layout.addWidget(self.tabela)
        self.setLayout(self.layout)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    janela = TabelaColorida()
    janela.show()
    sys.exit(app.exec_())
