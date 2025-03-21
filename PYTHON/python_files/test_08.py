import os
import sys

import win32print  # Apenas para Windows
from PyQt5.QtWidgets import (
    QApplication, QFileDialog, QMainWindow, QPushButton, QVBoxLayout, QWidget,
    QMessageBox, QDialog, QLabel, QComboBox, QCheckBox
)


class PrinterSelectionDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.printer_combo = None
        self.duplex_checkbox = None
        self.selected_printer = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Selecionar Impressora")
        self.setGeometry(100, 100, 400, 150)
        layout = QVBoxLayout()

        # Seleção da impressora
        label = QLabel("Selecione a impressora:")
        self.printer_combo = QComboBox()

        # Obtém lista de impressoras
        printers = self.get_printers()
        self.printer_combo.addItems(printers)

        # Define impressora padrão
        default_printer = self.get_default_printer()
        default_index = self.printer_combo.findText(default_printer)
        if default_index >= 0:
            self.printer_combo.setCurrentIndex(default_index)

        # Opção de impressão frente e verso
        self.duplex_checkbox = QCheckBox("Impressão frente e verso")
        self.duplex_checkbox.setChecked(True)

        # Botões
        print_button = QPushButton("Imprimir")
        cancel_button = QPushButton("Cancelar")
        print_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)

        # Adiciona widgets ao layout
        layout.addWidget(label)
        layout.addWidget(self.printer_combo)
        layout.addWidget(self.duplex_checkbox)
        layout.addWidget(print_button)
        layout.addWidget(cancel_button)
        self.setLayout(layout)

    def get_selected_printer(self):
        return self.printer_combo.currentText()

    def is_duplex_selected(self):
        return self.duplex_checkbox.isChecked()

    def get_printers(self):
        """Lista todas as impressoras disponíveis no sistema."""
        if sys.platform == "win32":
            return [printer[2] for printer in win32print.EnumPrinters(2)]
        else:  # Linux/macOS usa CUPS
            result = os.popen("lpstat -a").read()
            return [line.split()[0] for line in result.split("\n") if line]

    def get_default_printer(self):
        """Retorna a impressora padrão do sistema."""
        if sys.platform == "win32":
            return win32print.GetDefaultPrinter()
        else:  # Linux/macOS
            result = os.popen("lpstat -d").read()
            return result.split(": ")[1].strip() if ": " in result else ""


class PDFPrinter(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Impressão de PDF")
        self.setGeometry(100, 100, 400, 200)

        layout = QVBoxLayout()

        self.btn_select_pdf = QPushButton("Selecionar PDF para Impressão")
        self.btn_select_pdf.clicked.connect(self.select_and_print_pdf)

        layout.addWidget(self.btn_select_pdf)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def show_message(self, title, message, icon):
        msg = QMessageBox(self)
        msg.setIcon(icon)
        msg.setWindowTitle(title)
        msg.setText(message)
        msg.exec_()

    def select_and_print_pdf(self):
        pdf_file, _ = QFileDialog.getOpenFileName(self, "Selecione o PDF", "", "PDF Files (*.pdf)")

        if not pdf_file:
            self.show_message("Aviso", "Nenhum arquivo selecionado!", QMessageBox.Warning)
            return

        dialog = PrinterSelectionDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            selected_printer = dialog.get_selected_printer()
            is_duplex = dialog.is_duplex_selected()
            self.print_pdf(pdf_file, selected_printer, is_duplex)

    def print_pdf(self, pdf_path, printer_name, is_duplex):
        """Envia o arquivo PDF para a impressora selecionada"""
        try:
            if sys.platform == "win32":
                command = f'print /D:"{printer_name}" "{pdf_path}"'
            else:  # Linux/macOS
                duplex_option = "-o sides=two-sided-long-edge" if is_duplex else ""
                command = f'lp -d "{printer_name}" {duplex_option} "{pdf_path}"'

            resultado = os.system(command)

            if resultado == 0:
                self.show_message("Sucesso", f"Arquivo enviado para a impressora {printer_name}!", QMessageBox.Information)
            else:
                self.show_message("Erro", f"Falha ao imprimir na impressora {printer_name}!", QMessageBox.Critical)

        except Exception as e:
            self.show_message("Erro", f"Ocorreu um erro: {str(e)}", QMessageBox.Critical)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFPrinter()
    window.show()
    sys.exit(app.exec_())
