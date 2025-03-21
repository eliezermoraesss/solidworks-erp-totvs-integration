import os
import pandas as pd
from PyQt5.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QPushButton, QFileDialog, QTableWidget, QTableWidgetItem, QLabel, QProgressDialog
from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QDesktopServices
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

class PrintOP(QMainWindow):
    def __init__(self, dataframe):
        super().__init__()
        self.dataframe = dataframe
        self.setWindowTitle("Imprimir OP")
        self.setGeometry(300, 200, 800, 600)

        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.table = QTableWidget()
        self.load_data()
        layout.addWidget(self.table)

        self.print_btn = QPushButton("Imprimir")
        self.print_btn.clicked.connect(self.print_op)
        layout.addWidget(self.print_btn)

        self.close_btn = QPushButton("Fechar")
        self.close_btn.clicked.connect(self.close)
        layout.addWidget(self.close_btn)

        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)

    def load_data(self):
        self.table.setRowCount(len(self.dataframe))
        self.table.setColumnCount(len(self.dataframe.columns))
        self.table.setHorizontalHeaderLabels(self.dataframe.columns)

        for i, row in self.dataframe.iterrows():
            for j, cell in enumerate(row):
                self.table.setItem(i, j, QTableWidgetItem(str(cell)))

    def print_op(self):
        save_path = "\\\\192.175.175.4\\dados\\EMPRESA\\PRODUCAO\\ORDEM_DE_PRODUCAO"

        file_name, _ = QFileDialog.getSaveFileName(self, "Salvar PDF", save_path, "PDF Files (*.pdf)")
        if file_name:
            self.create_pdf(file_name)
            QDesktopServices.openUrl(QUrl.fromLocalFile(file_name))

    def create_pdf(self, file_path):
        progress = QProgressDialog("Publicando Ordem de Produção...", "Cancelar", 0, 100, self)
        progress.setWindowModality(Qt.WindowModal)
        progress.show()

        pdf = canvas.Canvas(file_path, pagesize=A4)
        pdf.setTitle("Ordem de Produção")

        logo_path = "empresa_logo.png"
        if os.path.exists(logo_path):
            pdf.drawImage(ImageReader(logo_path), 50, 750, width=100, height=50)

        pdf.setFont("Helvetica-Bold", 16)
        pdf.drawString(200, 780, "Ordem de Produção")

        progress.setValue(30)

        y_position = 700
        for _, row in self.dataframe.iterrows():
            pdf.drawString(50, y_position, f"OP: {row['OP']} - {row['Código Pai']} - {row['Descrição']} - {row['Quantidade']}")
            y_position -= 20
            if y_position < 100:
                pdf.showPage()
                y_position = 700

        progress.setValue(70)

        pdf.showPage()
        pdf.save()

        progress.setValue(100)
        progress.close()

if __name__ == "__main__":
    import sys
    from PyQt5.QtWidgets import QApplication

    df = pd.DataFrame({
        "OP": [f"{i:03d}" for i in range(1, 10001)],
        "Código Pai": [f"A{i:03d}" for i in range(1, 10001)],
        "Descrição": [f"Produto {chr(88 + (i % 3))}" for i in range(1, 10001)],
        "Quantidade": [i * 10 for i in range(1, 10001)]
    })

    app = QApplication(sys.argv)
    window = PrintOP(df)
    window.show()
    sys.exit(app.exec_())
