import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QMenu, QAction, QMessageBox
from PyQt5.QtCore import Qt, QPoint

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Tabela com Menu de Contexto')
        self.setGeometry(300, 300, 600, 400)

        self.table = QTableWidget(5, 3, self)
        self.setCentralWidget(self.table)
        self.table.setHorizontalHeaderLabels(['Coluna 1', 'Coluna 2', 'Coluna 3'])

        # Preenche a tabela com alguns dados de exemplo
        for row in range(5):
            for col in range(3):
                item = QTableWidgetItem(f'Item {row},{col}')
                self.table.setItem(row, col, item)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.showContextMenu)

    def showContextMenu(self, position):
        indexes = self.table.selectedIndexes()
        if indexes:
            menu = QMenu()
            custom_action = QAction('Ação Customizada', self)
            custom_action.triggered.connect(self.customAction)
            menu.addAction(custom_action)

            menu.exec_(self.table.viewport().mapToGlobal(position))

    def customAction(self):
        indexes = self.table.selectedIndexes()
        if indexes:
            row = indexes[0].row()
            col = indexes[0].column()
            item = self.table.item(row, col).text()
            QMessageBox.information(self, 'Ação Executada', f'Ação executada no item: {item}')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWin = MainWindow()
    mainWin.show()
    sys.exit(app.exec_())
