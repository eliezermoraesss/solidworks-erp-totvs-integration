import subprocess
import sys

from PyQt5.QtPrintSupport import QPrinterInfo
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QLineEdit, QPushButton, QFileDialog,
                             QMessageBox, QComboBox)


class PdfPrinterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Impressora de PDF")
        self.setGeometry(100, 100, 600, 200)

        # Configurar caminho do SumatraPDF (ajuste conforme necessário)
        self.sumatra_path = r"\\192.175.175.4\desenvolvimento\REPOSITORIOS\resources\SumatraPDF\SumatraPDF.exe"

        # Variáveis de configuração
        self.pdf_path = ""
        self.selected_printer = ""

        # Widgets
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout()

        # Selecionar PDF
        self.btn_pdf = QPushButton("Selecionar PDF")
        self.btn_pdf.clicked.connect(self.select_pdf)
        self.layout.addWidget(self.btn_pdf)

        self.lbl_pdf = QLineEdit()
        self.lbl_pdf.setReadOnly(True)
        self.layout.addWidget(self.lbl_pdf)


        # Selecionar Impressora
        self.cmb_printers = QComboBox()
        self.populate_printers()
        self.layout.addWidget(self.cmb_printers)

        # Botão Imprimir
        self.btn_print = QPushButton("Imprimir")
        self.btn_print.clicked.connect(self.print_pdf)
        self.layout.addWidget(self.btn_print)

        self.central_widget.setLayout(self.layout)

    def populate_printers(self):
        printers = QPrinterInfo.availablePrinters()
        for printer in printers:
            self.cmb_printers.addItem(printer.printerName())

    def select_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Selecionar PDF", "", "PDF Files (*.pdf)")
        if file_path:
            self.pdf_path = file_path
            self.lbl_pdf.setText(file_path)

    def print_pdf(self):
        if not self.pdf_path:
            QMessageBox.warning(self, "Erro", "Selecione um arquivo PDF!")
            return

        printer_name = self.cmb_printers.currentText()
        if not printer_name:
            QMessageBox.warning(self, "Erro", "Selecione uma impressora!")
            return

        # Configurações de impressão
        settings = "paper=A4,duplex,portrait"

        # Montar comando
        cmd = [
            self.sumatra_path,
            "-print-to", printer_name,
            "-print-settings", settings,
            self.pdf_path
        ]

        try:
            subprocess.run(cmd, check=True, capture_output=True, text=True)
            QMessageBox.information(self, "Sucesso", "PDF enviado para impressão com sucesso!")
        except subprocess.CalledProcessError as e:
            error_msg = f"Erro ao imprimir:\n{e.stderr}"
            QMessageBox.critical(self, "Erro", error_msg)
        except Exception as e:
            QMessageBox.critical(self, "Erro", str(e))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PdfPrinterApp()
    window.show()
    sys.exit(app.exec_())
