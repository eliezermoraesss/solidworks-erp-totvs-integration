import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableView, QVBoxLayout, QWidget, QStyledItemDelegate
from PyQt5.QtGui import QColor, QPainter
from PyQt5.QtCore import Qt, QAbstractTableModel

class MyTableModel(QAbstractTableModel):
    def __init__(self, data):
        super(MyTableModel, self).__init__()
        self._data = data

    def data(self, index, role):
        if role == Qt.DisplayRole:
            return self._data[index.row()][index.column()]

    def rowCount(self, index):
        return len(self._data)

    def columnCount(self, index):
        return len(self._data[0])

class BarDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        value = index.data()
        painter.save()

        # Draw background
        painter.setBrush(QColor(255, 255, 255))
        painter.drawRect(option.rect)

        # Draw bar
        bar_width = int((value / 100.0) * option.rect.width())  # Assuming max value is 100 for simplicity
        if value < 0:
            color = QColor(255, 0, 0)  # Red for negative values
        else:
            color = QColor(0, 255, 0)  # Green for positive values
        painter.setBrush(color)
        painter.drawRect(option.rect.x(), option.rect.y(), bar_width, option.rect.height())

        painter.restore()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("QTableView with Bar Delegate")
        self.setGeometry(100, 100, 600, 400)

        layout = QVBoxLayout()
        self.table = QTableView()

        # Sample data: [(day difference, other columns...)]
        data = [[-10], [5], [20], [-30], [15], [0]]
        self.model = MyTableModel(data)
        self.table.setModel(self.model)

        # Set the delegate for the first column
        delegate = BarDelegate(self.table)
        self.table.setItemDelegateForColumn(0, delegate)

        layout.addWidget(self.table)
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
