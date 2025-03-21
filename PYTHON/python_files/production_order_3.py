import os
import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QTableWidget, QVBoxLayout, QHBoxLayout,
    QWidget, QTableWidgetItem, QHeaderView, QLabel, QFileDialog, QMessageBox, QProgressDialog
)
from PyQt5.QtGui import QPixmap, QPainter
from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtPrintSupport import QPrinter
import barcode
from barcode.writer import ImageWriter
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

class PrintOP(QMainWindow):
    def __init__(self, dataframe):
        super().__init__()

        self.dataframe = dataframe
        self.setWindowTitle("Imprimir Ordem de Produção")
        self.setGeometry(100, 100, 900, 600)

        # Layout Principal
        main_layout = QHBoxLayout()

        # Container Esquerdo - Tabela
        self.table = QTableWidget()
        self.table.setRowCount(len(dataframe))
        self.table.setColumnCount(len(dataframe.columns))
        self.table.setHorizontalHeaderLabels(dataframe.columns)

        for i in range(len(dataframe)):
            for j in range(len(dataframe.columns)):
                self.table.setItem(i, j, QTableWidgetItem(str(dataframe.iloc[i, j])))

        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        left_layout = QVBoxLayout()
        left_layout.addWidget(QLabel("Selecione as Ordens de Produção"))
        left_layout.addWidget(self.table)

        left_container = QWidget()
        left_container.setLayout(left_layout)

        # Container Direito - Opções e Botões
        self.print_button = QPushButton("Imprimir")
        self.close_button = QPushButton("Fechar")
        self.print_button.clicked.connect(self.print_op)
        self.close_button.clicked.connect(self.close)

        right_layout = QVBoxLayout()
        right_layout.addWidget(QLabel("Opções de Impressão"))
        right_layout.addWidget(self.print_button)
        right_layout.addWidget(self.close_button)
        right_layout.addStretch()

        right_container = QWidget()
        right_container.setLayout(right_layout)

        # Adicionando Containers ao Layout Principal
        main_layout.addWidget(left_container, 2)
        main_layout.addWidget(right_container, 1)

        main_widget = QWidget()
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

    def print_op(self):
        save_path = r"\\192.175.175.4\dados\EMPRESA\PRODUCAO\ORDEM_DE_PRODUCAO"
        if not os.path.exists(save_path):
            os.makedirs(save_path)

        file_path = os.path.join(save_path, "ordem_de_producao.pdf")

        progress = QProgressDialog("Publicando Ordem de Produção...", None, 0, 100, self)
        progress.setWindowModality(Qt.WindowModal)
        progress.show()

        printer = QPrinter()
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(file_path)

        painter = QPainter(printer)

        # Cabeçalho
        logo_path = "logo_empresa.png"
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path).scaled(100, 50, Qt.KeepAspectRatio)
            painter.drawPixmap(30, 30, pixmap)

        barcode_data = "123456789"
        barcode_path = self.create_barcode(barcode_data)

        if barcode_path:
            barcode_img = QPixmap(barcode_path).scaled(100, 50, Qt.KeepAspectRatio)
            painter.drawPixmap(500, 30, barcode_img)

        # Tabela Hierárquica
        y_offset = 100
        for row in range(len(self.dataframe)):
            line = " | ".join([str(self.dataframe.iloc[row, col]) for col in range(len(self.dataframe.columns))])
            painter.drawText(30, y_offset, line)
            y_offset += 20

        painter.end()

        progress.setValue(100)
        QMessageBox.information(self, "Impressão Concluída", "A Ordem de Produção foi gerada com sucesso!")
        QUrl.fromLocalFile(file_path).toString()

    def create_barcode(self, data):
        barcode_path = "barcode.png"
        ean = barcode.get('code128', data, writer=ImageWriter())
        ean.save(barcode_path)
        return barcode_path

if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Exemplo de DataFrame
    data = {
        "OP": [str(i) for i in range(1, 10001)],
        "Código Pai": [f"A{i:03d}" for i in range(1, 10001)],
        "Descrição": [f"Produto {i}" for i in range(1, 10001)],
        "Quantidade": [i * 10 for i in range(1, 10001)]
    }
    df = pd.DataFrame(data)

    window = PrintOP(df)
    window.show()
    sys.exit(app.exec_())
