import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, QPushButton
from PyQt5.QtCore import Qt  # Adicione esta linha

class ExpandableTable(QMainWindow):
    def __init__(self):
        super(ExpandableTable, self).__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle('Tabela Expansível - PyQt5')
        self.setGeometry(100, 100, 600, 400)

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout()

        self.table = QTableWidget(self)
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(['ID', 'Item', ''])

        self.addTableRow(1, 'Item 1', 'Conteúdo Expansível 1')
        self.addTableRow(2, 'Item 2', 'Conteúdo Expansível 2')

        layout.addWidget(self.table)

        central_widget.setLayout(layout)

    def addTableRow(self, id, item, content):
        rowPosition = self.table.rowCount()
        self.table.insertRow(rowPosition)

        id_item = QTableWidgetItem(str(id))
        item_item = QTableWidgetItem(item)

        expand_button = QPushButton('+')
        expand_button.clicked.connect(lambda: self.toggleRowContent(rowPosition))

        self.table.setItem(rowPosition, 0, id_item)
        self.table.setItem(rowPosition, 1, item_item)
        self.table.setCellWidget(rowPosition, 2, expand_button)

        # Use QTableWidgetItem for content
        content_item = QTableWidgetItem(content)
        content_item.setFlags(content_item.flags() & ~Qt.ItemIsEditable)  # Make the content cell non-editable
        content_item.setBackground(self.table.palette().alternateBase())  # Match background color
        content_item.setHidden(True)  # This line is not needed

        self.table.setItem(rowPosition + 1, 0, content_item)
        self.table.setSpan(rowPosition, 2, 2, 1)

    def toggleRowContent(self, row):
        content_item = self.table.item(row + 1, 0)
        if content_item.flags() & Qt.ItemIsEnabled:  # Check if the item is enabled
            content_item.setFlags(content_item.flags() & ~Qt.ItemIsEnabled)  # Disable the item
        else:
            content_item.setFlags(content_item.flags() | Qt.ItemIsEnabled)  # Enable the item


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ExpandableTable()
    window.show()
    sys.exit(app.exec_())
