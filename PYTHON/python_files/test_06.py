from PyQt5.QtPrintSupport import QPrinter, QPrintDialog
from PyQt5.QtWidgets import QApplication, QFileDialog, QMessageBox
from PyQt5.QtGui import QPainter
import sys
from PyPDF2 import PdfReader

def print_pdf():
    app = QApplication(sys.argv)

    file_dialog = QFileDialog()
    file_path, _ = file_dialog.getOpenFileName(None, "Selecione o PDF", "", "PDF Files (*.pdf)")

    if not file_path:
        return

    printer = QPrinter(QPrinter.HighResolution)
    printer.setPageSize(QPrinter.A4)
    printer.setOutputFormat(QPrinter.NativeFormat)
    printer.setDuplex(QPrinter.DuplexAuto)

    print_dialog = QPrintDialog(printer)
    if print_dialog.exec_() == QPrintDialog.Accepted:
        reader = PdfReader(file_path)
        painter = QPainter(printer)

        for i, page in enumerate(reader.pages):
            # Renderiza a página diretamente no dispositivo de impressão
            page.render(painter)

            if i < len(reader.pages) - 1:
                printer.newPage()

        painter.end()
        QMessageBox.information(None, "Sucesso", "O PDF foi enviado para a impressora!")

if __name__ == "__main__":
    print_pdf()