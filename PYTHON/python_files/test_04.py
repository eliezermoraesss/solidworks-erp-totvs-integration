import PyPDF2
from PyQt5.QtWidgets import QFileDialog, QApplication, QMessageBox
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog
from PyQt5.QtGui import QPainter
from PyQt5.QtCore import QRectF



def imprimir_pdf_pypdf2(caminho_pdf):
    try:
        with open(caminho_pdf, "rb") as arquivo_pdf:
            leitor_pdf = PyPDF2.PdfReader(arquivo_pdf)
            num_paginas = len(leitor_pdf.pages)

            app = QApplication([])
            printer = QPrinter()
            dialog = QPrintDialog(printer)

            printer.setPageSize(QPrinter.A4)  # Define o tamanho da página como A4
            printer.setDuplex(QPrinter.DuplexLongSide)  # Define a impressão frente e verso

            if dialog.exec_() == QPrintDialog.Accepted:
                painter = QPainter(printer)
                page_size = printer.pageRect().size()

                for page_num in range(num_paginas):
                    page = leitor_pdf.pages[page_num]
                    page_rect = page.mediabox
                    scale = min(page_size.width() / page_rect.width(), page_size.height() / page_rect.height())
                    painter.setViewport(page_rect)
                    painter.setWindow(page_rect)
                    painter.scale(scale, scale)
                    page.extract_text(page_rect, painter)

                    if page_num < num_paginas - 1:
                        printer.newPage()

                painter.end()
            app.exec_()

    except FileNotFoundError:
        QMessageBox.critical(None, "Erro", "Arquivo PDF não encontrado.")
    except Exception as e:
        QMessageBox.critical(None, "Erro", f"Ocorreu um erro ao imprimir o PDF: {e}")

# Exemplo de uso:
if __name__ == "__main__":
    # caminho_pdf = QFileDialog.getOpenFileName(None, "Selecionar Arquivo PDF", "", "Arquivos PDF (*.pdf)")[0]
    caminho_pdf = r'V:\ORDEM_DE_PRODUCAO\OP_01483601022_M-039-033-110.pdf'
    if caminho_pdf:
        imprimir_pdf_pypdf2(caminho_pdf)