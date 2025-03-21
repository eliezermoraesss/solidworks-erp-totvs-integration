import os
from datetime import datetime
from typing import List, Dict

import pandas as pd
from PyQt5 import QtWidgets
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import (QHBoxLayout, QVBoxLayout, QTableWidget,
                             QTableWidgetItem, QPushButton, QHeaderView, QCheckBox, QMessageBox)
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph


class PDFGeneratorThread(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, data: List[Dict], output_path: str):
        super().__init__()
        self.data = data
        self.output_path = output_path

    def run(self):
        try:
            generate_production_order_pdf(self.data, self.output_path)
            self.finished.emit(self.output_path)
        except Exception as e:
            self.error.emit(str(e))

def generate_qr_code(data: str, size: int = 10) -> str:
    """Gera um código QR e retorna o caminho para a imagem"""
    qr = qrcode.QRCode(version=1, box_size=size, border=5)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")

    script_dir = os.path.dirname(os.path.abspath(__file__))
    temp_qr_path = os.path.join(script_dir, '..', 'resources', 'images', 'barcode.png')
    return temp_qr_path

def get_company_logo_path():
    """Retorna o caminho do logo da empresa"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(script_dir, '..', 'resources', 'images', 'logo_enaplic.jpg')

def generate_production_order_pdf(data: List[Dict], output_path: str):
    """Gera PDF da Ordem de Produção"""
    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4

    # Margens reduzidas
    margin = 15 * mm

    # Cabeçalho
    logo_path = get_company_logo_path()
    c.drawImage(logo_path, margin, height - margin * 2, width=50, preserveAspectRatio=True)

    # QR Code / Código de Barras
    qr_path = generate_qr_code(str(data[0]['OP']))
    c.drawImage(qr_path, width - margin * 4, height - margin * 2, width=50, preserveAspectRatio=True)

    # Informações da Ordem de Produção
    styles = getSampleStyleSheet()
    op_info = [
        f"OP: {data[0]['OP']}",
        f"Produto: {data[0]['Código']} - {data[0]['Descrição']}",
        f"Quantidade: {data[0]['Quantidade']}",
        f"Data Abertura: {data[0]['Data Abertura']}",
        f"Previsão Entrega: {data[0]['Prev. Entrega']}"
    ]

    y_position = height - margin * 4
    for line in op_info:
        p = Paragraph(line, styles['Normal'])
        p.wrapOn(c, width - 2*margin, 20)
        p.drawOn(c, margin, y_position)
        y_position -= 15

    # Limpar recursos temporários
    if os.path.exists(qr_path):
        os.remove(qr_path)

    c.showPage()
    c.save()

class PrintProductionOrderDialog(QtWidgets.QDialog):
    def __init__(self, dataframe: pd.DataFrame, parent=None):
        super().__init__(parent)
        self.pdf_thread = None
        self.dataframe = dataframe
        self.print_production_order()

    def print_production_order(self):
        # Lógica de impressão da ordem de produção
        selected_rows = set(index.row() for index in self.results_table.selectedIndexes())

        if not selected_rows:
            # Lógica para lidar com nenhuma seleção
            return

        # Preparar dados para impressão
        selected_data = self.dataframe.iloc[list(selected_rows)].to_dict('records')

        # Definir caminho de saída
        output_dir = r"\\192.175.175.4\dados\EMPRESA\PRODUCAO\ORDEM_DE_PRODUCAO"
        os.makedirs(output_dir, exist_ok=True)

        output_path = os.path.join(output_dir, f"OP_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")

        # Iniciar thread de geração de PDF
        self.pdf_thread = PDFGeneratorThread(selected_data, output_path)
        self.pdf_thread.finished.connect(self.on_pdf_generation_complete)
        self.pdf_thread.error.connect(self.on_pdf_generation_error)
        self.pdf_thread.start()

    def on_pdf_generation_complete(self, path):
        # Lógica após geração do PDF (abrir, mostrar mensagem)
        QMessageBox.information(self, "PDF Gerado", f"PDF gerado em: {path}")

    def on_pdf_generation_error(self, error):
        # Tratar erros na geração do PDF
        QMessageBox.critical(self, "Erro na Geração do PDF", f"Erro na geração do PDF: {error}")
