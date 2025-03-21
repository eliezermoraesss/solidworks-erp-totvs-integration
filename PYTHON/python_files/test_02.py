import sys

import fitz  # PyMuPDF
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPainter, QImage
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog
from PyQt5.QtWidgets import QApplication, QMessageBox


def print_pdf(pdf_path):
    app = QApplication(sys.argv)

    printer = QPrinter(QPrinter.HighResolution)
    printer.setPageSize(QPrinter.A4)
    printer.setOutputFormat(QPrinter.NativeFormat)
    printer.setDuplex(QPrinter.DuplexAuto)

    dialog = QPrintDialog(printer)
    if dialog.exec_() == QPrintDialog.Accepted:
        document = fitz.open(pdf_path)
        painter = QPainter(printer)

        for page_number in range(document.page_count):
            page = document.load_page(page_number)
            zoom = 2.0
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)

            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)

            rect = painter.viewport()
            scaled_size = img.size().scaled(rect.size(), Qt.KeepAspectRatio)
            painter.setViewport(rect.x(), rect.y(), scaled_size.width(), scaled_size.height())
            painter.setWindow(img.rect())

            painter.drawImage(0, 0, img)

            if page_number < document.page_count - 1:
                printer.newPage()

        painter.end()
        QMessageBox.information(None, "Sucesso", "O PDF foi enviado para a impressora!")

if __name__ == "__main__":
    print_pdf('V:\\ORDEM_DE_PRODUCAO\\OP_01483601022_M-039-033-110.pdf')
