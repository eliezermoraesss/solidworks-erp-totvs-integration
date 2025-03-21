from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QFileDialog
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog
from PyQt6.QtPdf import QPdfDocument
from PyQt6.QtCore import Qt, QMarginsF
import sys

class PDFPrinter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Impressão de PDF")
        self.setGeometry(100, 100, 400, 200)

        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Layout
        layout = QVBoxLayout(central_widget)

        # Botão para selecionar e imprimir PDF
        self.print_button = QPushButton("Selecionar e Imprimir PDF")
        self.print_button.clicked.connect(self.print_pdf)
        layout.addWidget(self.print_button)

    def print_pdf(self):
        # Diálogo para selecionar arquivo
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar PDF",
            "",
            "Arquivos PDF (*.pdf)"
        )

        if file_name:
            # Configurar impressora
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)

            # Configurações padrão
            printer.setPageSize(QPrinter.PageSize.A4)
            printer.setDuplex(QPrinter.DuplexMode.DuplexLongSide)  # Impressão frente e verso
            printer.setPageMargins(QMarginsF(10, 10, 10, 10))  # Margens em mm

            # Abrir diálogo de impressão
            dialog = QPrintDialog(printer, self)
            if dialog.exec() == QPrintDialog.DialogCode.Accepted:
                # Carregar o documento PDF
                document = QPdfDocument()
                document.load(file_name)

                if document.status() == QPdfDocument.Status.Ready:
                    # Configurar orientação baseada na primeira página
                    first_page = document.pagePointSize(0)
                    if first_page.width() > first_page.height():
                        printer.setPageOrientation(Qt.Orientation.Landscape)
                    else:
                        printer.setPageOrientation(Qt.Orientation.Portrait)

                    # Imprimir documento
                    try:
                        success = document.print(printer)
                        if success:
                            print("Documento impresso com sucesso!")
                        else:
                            print("Erro ao imprimir documento")
                    except Exception as e:
                        print(f"Erro durante a impressão: {str(e)}")
                else:
                    print("Erro ao carregar o PDF")

def main():
    app = QApplication(sys.argv)
    window = PDFPrinter()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()