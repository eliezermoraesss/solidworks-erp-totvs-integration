from PyQt5.QtWidgets import (QMainWindow, QPushButton, QFileDialog, QMessageBox,
                             QVBoxLayout, QWidget, QLabel, QScrollArea)
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog, QPrintPreviewDialog
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap, QPainter, QImage
import sys
from PyQt5.QtWidgets import QApplication
import PyPDF2
from PIL import Image
import fitz  # PyMuPDF
import io

class PDFPrinter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_file = None
        self.pdf_document = None
        self.current_page = 0
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Visualizador e Impressão de PDF')
        self.setGeometry(100, 100, 800, 600)

        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Layout principal
        layout = QVBoxLayout(central_widget)

        # Área de visualização
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.preview_label = QLabel()
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.scroll_area.setWidget(self.preview_label)

        # Botões
        button_layout = QVBoxLayout()
        self.open_button = QPushButton('Abrir PDF', self)
        self.open_button.clicked.connect(self.openPDF)

        self.print_button = QPushButton('Imprimir', self)
        self.print_button.clicked.connect(self.printPDF)
        self.print_button.setEnabled(False)

        self.preview_print_button = QPushButton('Pré-visualizar Impressão', self)
        self.preview_print_button.clicked.connect(self.previewPrint)
        self.preview_print_button.setEnabled(False)

        # Adiciona widgets ao layout
        button_layout.addWidget(self.open_button)
        button_layout.addWidget(self.preview_print_button)
        button_layout.addWidget(self.print_button)

        layout.addWidget(self.scroll_area, stretch=1)
        layout.addLayout(button_layout)

    def openPDF(self):
        fileName, _ = QFileDialog.getOpenFileName(
            self,
            "Selecione o PDF",
            "",
            "Arquivos PDF (*.pdf)"
        )

        if fileName:
            try:
                self.current_file = fileName
                # Abre o PDF usando PyMuPDF
                self.pdf_document = fitz.open(fileName)

                # Habilita botões
                self.print_button.setEnabled(True)
                self.preview_print_button.setEnabled(True)

                # Mostra primeira página como preview
                self.displayPreview(0)

                QMessageBox.information(
                    self,
                    "Sucesso",
                    "PDF carregado com sucesso!"
                )

            except Exception as e:
                QMessageBox.critical(
                    self,
                    "Erro",
                    f"Erro ao abrir o PDF: {str(e)}"
                )

    def displayPreview(self, page_number):
        if self.pdf_document and page_number < len(self.pdf_document):
            try:
                # Obtém a página do PDF
                page = self.pdf_document[page_number]

                # Renderiza a página como uma imagem
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # zoom 2x para melhor qualidade

                # Converte para QImage
                img_data = pix.samples
                qimg = QImage(img_data, pix.width, pix.height, pix.stride, QImage.Format_RGB888)

                # Converte para QPixmap e exibe
                pixmap = QPixmap.fromImage(qimg)
                scaled_pixmap = pixmap.scaled(
                    self.scroll_area.size(),
                    Qt.KeepAspectRatio,
                    Qt.SmoothTransformation
                )

                self.preview_label.setPixmap(scaled_pixmap)

            except Exception as e:
                QMessageBox.critical(
                    self,
                    "Erro",
                    f"Erro ao exibir página: {str(e)}"
                )

    def previewPrint(self):
        if not self.current_file:
            return

        printer = QPrinter(QPrinter.HighResolution)
        printer.setPageSize(QPrinter.A4)
        printer.setDuplex(QPrinter.DuplexLongSide)

        preview = QPrintPreviewDialog(printer, self)
        preview.paintRequested.connect(self.paintPreview)
        preview.exec_()

    def paintPreview(self, printer):
        if self.pdf_document:
            painter = QPainter(printer)
            for page_num in range(len(self.pdf_document)):
                if page_num > 0:
                    printer.newPage()

                page = self.pdf_document[page_num]
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))

                img_data = pix.samples
                qimg = QImage(img_data, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
                pixmap = QPixmap.fromImage(qimg)

                # Ajusta o tamanho da página
                printer_rect = printer.pageRect(QPrinter.DevicePixel)
                pixmap = pixmap.scaled(
                    printer_rect.size(),
                    Qt.KeepAspectRatio,
                    Qt.SmoothTransformation
                )

                # Centraliza a imagem na página
                x = (printer_rect.width() - pixmap.width()) / 2
                y = (printer_rect.height() - pixmap.height()) / 2

                painter.drawPixmap(int(x), int(y), pixmap)

            painter.end()

    def printPDF(self):
        if not self.current_file:
            return

        try:
            printer = QPrinter(QPrinter.HighResolution)
            printer.setPageSize(QPrinter.A4)
            printer.setDuplex(QPrinter.DuplexLongSide)

            dialog = QPrintDialog(printer, self)

            if dialog.exec_() == QPrintDialog.Accepted:
                self.paintPreview(printer)
                QMessageBox.information(
                    self,
                    "Sucesso",
                    "Documento enviado para impressão!"
                )

        except Exception as e:
            QMessageBox.critical(
                self,
                "Erro",
                f"Erro ao imprimir: {str(e)}"
            )

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = PDFPrinter()
    ex.show()
    sys.exit(app.exec_())
