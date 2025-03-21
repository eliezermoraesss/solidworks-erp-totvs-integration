import os
import subprocess
import tempfile
from datetime import datetime

import pandas as pd
import win32api
import win32print
from PyPDF2 import PdfMerger
from PyQt5 import QtWidgets
from PyQt5.QtCore import QThread, pyqtSignal, QUrl
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtWidgets import (QVBoxLayout, QMessageBox, QProgressBar, QLabel, QDialog, QComboBox, QCheckBox, QPushButton)
from barcode import Code128
from barcode.writer import ImageWriter
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.fonts import addMapping
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle, Paragraph
from sqlalchemy import create_engine

from src.app.utils.db_mssql import setup_mssql
from src.app.utils.utils import exibir_mensagem
from src.dialog.information_dialog import information_dialog


class PDFGeneratorThread(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    progress = pyqtSignal(int, int, int)

    def __init__(self, df: pd.DataFrame, df_op_table, output_dir: str, selected_row):
        super().__init__()
        self.output_path = None
        self.df = df
        self.output_dir = output_dir
        self.df_op_table = df_op_table
        self.selected_row = selected_row

    def run(self):
        try:
            self.df.reset_index(drop=True, inplace=True)
            total_rows = len(self.df)
            for index, row in self.df.iterrows():
                progress = int((index + 1) / total_rows * 100)
                self.output_path = os.path.join(self.output_dir,
                                                f"OP_{row['OP'].strip()}_{row['Código'].strip()}.pdf")
                self.progress.emit(progress, index + 1, total_rows)
                generate_production_order_pdf(row, self.output_path, self.df_op_table, self.selected_row)
            self.finished.emit(self.output_path)
        except Exception as e:
            self.error.emit(str(e))


def get_resource_path(resource_type: str, filename: str) -> str:
    """Returns the path for various resource types"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(script_dir, '..', '..', 'resources', resource_type, filename)


def create_header_style():
    """Creates and returns custom header style"""
    return ParagraphStyle(
        'CustomHeader',
        parent=getSampleStyleSheet()['Heading3'],
        fontSize=10,
        spaceAfter=30,
        alignment=TA_LEFT,
        fontName='Courier-New'
    )


def generate_hierarchical_table(df: pd.DataFrame, canvas_obj, y_position: float):
    """Generates the hierarchical table for the production order"""
    # Define column widths and headers
    col_widths = [70, 70, 336, 60]  # Adjusted widths
    headers = ['OP Pai', 'Código Pai', 'Descrição', 'Quantidade']

    # Prepare data for table
    table_data = [headers]
    for _, row in df.iterrows():
        descricao = str(row['Descrição']).strip()
        if len(descricao) > 60:
            descricao = descricao[:60] + '...'
        table_data.append([
            str(row['OP']),
            str(row['Código Pai']).strip(),
            descricao,
            str(row['Quantidade'])
        ])

    # Create and style table
    table = Table(table_data, colWidths=col_widths)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.transparent),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Courier-New-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('FONTNAME', (0, 1), (-1, -1), 'Courier-New'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.transparent),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))

    # Draw table
    table.wrapOn(canvas_obj, 530, 400)
    table_height = table._height

    # Check if the table fits on the current page
    if y_position - table_height < 0:
        canvas_obj.showPage()  # Create a new page
        y_position = 800  # Reset y_position for the new page

    # Draw table
    table.drawOn(canvas_obj, 30, y_position - table_height)


def consultar_hierarquia_tabela(codigo):
    username, password, database, server = setup_mssql()
    driver = '{SQL Server}'
    conn_str = (f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};'
                f'PWD={password}')
    engine = create_engine(f'mssql+pyodbc:///?odbc_connect={conn_str}')

    query_onde_usado = f"""
                    SELECT
                        STRUT.G1_COD AS "Código",
                        PROD.B1_DESC "Descrição",
                        STRUT.G1_QUANT AS "Quantidade",
                        STRUT.G1_XUM AS "Unid"
                    FROM
                        {database}.dbo.SG1010 STRUT
                    INNER JOIN
                        {database}.dbo.SB1010 PROD
                    ON
                        STRUT.G1_COD = PROD.B1_COD
                    WHERE 
                        G1_COMP = '{codigo}'
                    AND STRUT.G1_REVFIM = (
                        SELECT 
                            MAX(G1_REVFIM)
                        FROM 
                            {database}.dbo.SG1010
                        WHERE 
                            G1_COD = STRUT.G1_COD
                        AND 
                            G1_REVFIM <> 'ZZZ'
                        AND 
                            D_E_L_E_T_ <> '*')
                    AND 
                        STRUT.G1_REVFIM <> 'ZZZ'
                    AND 
                        STRUT.D_E_L_E_T_ <> '*'
                    AND 
                        PROD.D_E_L_E_T_ <> '*'
                    ORDER BY 
                        B1_DESC ASC;
                """
    try:
        with engine.connect() as connection:
            dataframe = pd.read_sql(query_onde_usado, connection)
        return dataframe if not dataframe.empty else pd.DataFrame()

    except Exception as ex:
        print(f"Erro ao consultar a hierarquia: {ex}")
        exibir_mensagem('Erro ao consultar itens pais da OP', f'Erro: {str(ex)}', 'error')
        return pd.DataFrame()

    finally:
        # Fecha a conexão com o banco de dados se estiver aberta
        engine.dispose()


def generate_barcode(data):
    temp_barcode = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
    Code128(data, writer=ImageWriter()).write(temp_barcode)
    temp_barcode.close()
    return temp_barcode.name


def generate_production_order_pdf(row: pd.Series, output_path: str, dataframe_geral, selected_row):
    data_hora_impressao = datetime.now().strftime('%d/%m/%Y   %H:%M:%S')
    codigo = row['Código'].strip()
    num_qp = row['QP/QR'].strip().zfill(6)
    num_op = row['OP'].strip()
    tipo = row['Tipo'].strip()
    op_geral = row['OP GERAL']
    op_aglutinada = row['Aglutinado']

    if selected_row:
        op_aglutinada = 'S' if op_aglutinada == 'Sim' else ''

    """Generates PDF for a single Production Order"""
    # Create PDF
    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4
    margin = 10 * mm

    # Logo
    logo_path = get_resource_path('images', 'logo_enaplic.jpg')
    logo_width = 80
    logo_x = margin
    logo_y = height - margin * 3
    c.drawImage(logo_path, logo_x, logo_y, width=logo_width, preserveAspectRatio=True)

    # Title
    title_text = f"ORDEM DE PRODUÇÃO - {num_op}"
    title_font_size = 12
    c.setFont("Courier-New-Bold", title_font_size)
    title_width = c.stringWidth(title_text, "Courier-New-Bold", title_font_size)
    title_x = (width - title_width) / 2
    title_y = 810  # Adjust this value as needed to position the title vertically
    c.drawString(title_x, title_y, title_text)

    title_text = f"{tipo}: {num_qp} {row['Projeto']}" if tipo == 'QP' else f"{row['Projeto']}"
    title_font_size = 12
    c.setFont("Courier-New-Bold", title_font_size)
    title_width = c.stringWidth(title_text, "Courier-New-Bold", title_font_size)
    title_x = (width - title_width) / 2
    title_y = title_y - 20  # Adjust this value as needed to position the title vertically
    c.drawString(title_x, title_y, title_text)

    # Add line below the title
    line_y = title_y - 10
    c.line(margin, line_y, width - margin, line_y)

    # Barcode
    barcode_width = 100
    barcode_x = 475
    barcode_y = 565  # Adjust this value as needed to position the barcode vertically
    barcode_path = generate_barcode(num_op)
    c.drawImage(barcode_path, barcode_x, barcode_y, width=barcode_width, preserveAspectRatio=True)

    # Informações da Ordem de Produção
    header_style = create_header_style()

    if not selected_row:
        data_emissao = datetime.strptime(row['Data Abertura'].strip(), "%Y%m%d").strftime("%d/%m/%Y")
        data_entrega = datetime.strptime(row['Prev Entrega'].strip(), "%Y%m%d").strftime("%d/%m/%Y")
    else:
        data_emissao = row['Data Abertura'].strip()
        data_entrega = row['Prev Entrega'].strip()

    op_info = [
        f"Produto: {row['Código'].strip()}  {row['Descrição'].strip()}",
        f"Quantidade: {row['Quantidade']}   {row['Unid']}",
        f"Centro de Custo: {row['Código CC']}   {row['Centro de Custo']}",
        f"Dt. Abertura da OP: {data_emissao}",
        f"Previsão de Entrega: {data_entrega}",
        f"Dt. Impressão da OP: {data_hora_impressao}",
        f"Observação: {row['Observação'].strip()}"
    ]

    y_position = 750  # Adjust this value as needed to position the information vertically
    for line in op_info:
        p = Paragraph(line, header_style)
        p.wrapOn(c, width - 2 * margin, 20)
        p.drawOn(c, margin, y_position)
        y_position -= 15

    # Add line below the title
    c.line(margin, y_position, width - margin, y_position)
    y_position -= 15

    # Tabela hierárquica
    dataframe_onde_usado = consultar_hierarquia_tabela(codigo)

    if not dataframe_onde_usado.empty:
        # Get the list of códigos from dataframe_onde_usado
        codigos_onde_usado = dataframe_onde_usado['Código'].unique()
        codigos_onde_usado = [item.strip() for item in codigos_onde_usado]

        if op_aglutinada == 'S':
            # Filter dataframe_geral to only include rows where 'Código' is in codigos_onde_usado and 'QP' matches num_qp
            dataframe_filtrado = dataframe_geral[
                dataframe_geral['Código'].str.strip().isin(codigos_onde_usado) &
                (dataframe_geral['QP/QR'].str.strip() == num_qp)
                ]
        else:
            dataframe_aglutinados = dataframe_geral[
                dataframe_geral['Código'].str.strip().isin(codigos_onde_usado) &
                (dataframe_geral['QP/QR'].str.strip() == num_qp) &
                (dataframe_geral['Aglutinado'].str.strip() == 'S')
                ]
            dataframe_nao_aglutinados = dataframe_geral[
                dataframe_geral['Código'].str.strip().isin(codigos_onde_usado) &
                (dataframe_geral['QP/QR'].str.strip() == num_qp) &
                (dataframe_geral['Aglutinado'].str.strip() != 'S') &
                (dataframe_geral['OP'].str.contains(op_geral, na=False))
                ]
            dataframe_filtrado = pd.concat([dataframe_aglutinados, dataframe_nao_aglutinados])

        # Add the OP column to the filtered dataframe if it doesn't exist
        if 'OP' not in dataframe_filtrado.columns:
            dataframe_filtrado['OP'] = row['OP']

        # Merge the filtered dataframe with onde_usado to get additional information
        dataframe_final = pd.merge(
            dataframe_filtrado,
            dataframe_onde_usado,
            on='Código',
            suffixes=('', '_onde_usado')
        )

        # Multiplica quantidade usada do filho pela quantidade do pai
        dataframe_final['Qtd Usada'] = dataframe_final['Quantidade'] * dataframe_final['Quantidade_onde_usado']

        # Select columns as needed
        dataframe_final = dataframe_final[[
            'OP',
            'Código',
            'Descrição',
            'Qtd Usada'
        ]].copy()

        # Rename columns to match the expected format
        dataframe_final = dataframe_final.rename(columns={
            'Código': 'Código Pai',
            'Qtd Usada': 'Quantidade'
        })

        # Sort the dataframe if needed
        dataframe_final = dataframe_final.sort_values('OP')

        # Title hierarchical table
        title_text = f"LISTA DOS PAIS (ONDE É USADO)"
        title_font_size = 10
        c.setFont("Courier-New-Bold", title_font_size)
        title_width = c.stringWidth(title_text, "Courier-New-Bold", title_font_size)
        title_x = (width - title_width) / 2
        c.drawString(title_x, y_position, title_text)
        y_position -= 15

        # Total destinado
        total_destinado = dataframe_final['Quantidade'].sum()
        p = Paragraph(f"Total destinado: {total_destinado}  {row['Unid']}", header_style)
        p.wrapOn(c, width - 2 * margin, 20)
        p.drawOn(c, margin, y_position)
        y_position -= 5

        # Generate the table
        table_y_position = y_position
        generate_hierarchical_table(dataframe_final, c, table_y_position)
    else:
        print(f"No hierarchical data found for código: {row['Código']}")

    # Roteiro
    workflow_path = get_resource_path('images', 'roteiro_v3.png')
    workflow_y_position = table_y_position - 780  # Adjust this value as needed to position the workflow vertically
    c.drawImage(workflow_path, margin, workflow_y_position, width=width - 2 * margin, preserveAspectRatio=True)

    # Save first page
    c.save()

    # Handle technical drawing (Page 2)
    codigo_desenho = row['Código'].strip()
    drawing_path = os.path.normpath(os.path.join(
        r"\\192.175.175.4\dados\EMPRESA\PROJETOS\PDF-OFICIAL",
        f"{codigo_desenho}.PDF"
    ))

    if os.path.exists(drawing_path):
        merger = PdfMerger()
        merger.append(output_path)  # First append the production order page
        merger.append(drawing_path)  # Then append the technical drawing

        # Create a temporary file for the merged result
        temp_output = output_path + '.temp'
        merger.write(temp_output)
        merger.close()

        # Replace the original file with the merged result
        os.replace(temp_output, output_path)
    else:
        # Create a new page with the "Drawing not found" message
        c = canvas.Canvas(output_path + '.temp', pagesize=A4)
        width, height = A4

        c.showPage()
        c.setFont("Courier-New-Italic", 24)
        c.drawCentredString(width / 2, height / 2, "DESENHO NÃO ENCONTRADO")
        c.save()

        # Merge the original page with the "Drawing not found" page
        merger = PdfMerger()
        merger.append(output_path)
        merger.append(output_path + '.temp')

        temp_output = output_path + '.merged'
        merger.write(temp_output)
        merger.close()

        # Clean up temporary files and replace the original
        os.remove(output_path + '.temp')
        os.replace(temp_output, output_path)

    # Open the generated PDF
    QDesktopServices.openUrl(QUrl(QUrl.fromLocalFile(output_path)))


def registrar_fonte_personalizada():
    # Registra todas as variações da fonte Courier
    pdfmetrics.registerFont(TTFont('Courier-New', r'C:\WINDOWS\FONTS\COUR.TTF'))
    pdfmetrics.registerFont(TTFont('Courier-New-Bold', r'C:\WINDOWS\FONTS\COURBD.TTF'))
    pdfmetrics.registerFont(TTFont('Courier-New-Italic', r'C:\WINDOWS\FONTS\COURI.TTF'))
    pdfmetrics.registerFont(TTFont('Courier-New-BoldItalic', r'C:\WINDOWS\FONTS\COURBI.TTF'))

    # Mapeia as variações da fonte
    addMapping('Courier-New', 0, 0, 'Courier-New')  # normal
    addMapping('Courier-New', 1, 0, 'Courier-New-Bold')  # bold
    addMapping('Courier-New', 0, 1, 'Courier-New-Italic')  # italic
    addMapping('Courier-New', 1, 1, 'Courier-New-BoldItalic')  # bold & italic


class PrinterSelectionDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.duplex_checkbox = None
        self.printer_combo = None
        self.selected_printer = None
        self.is_duplex = True
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Selecionar Impressora")
        self.setGeometry(100, 100, 400, 150)
        layout = QVBoxLayout()

        # Printer selection combo box
        label = QLabel("Selecione a impressora:")
        self.printer_combo = QComboBox()

        # Get list of printers
        printers = [printer[2] for printer in win32print.EnumPrinters(2)]  # 2 = Installed printers
        self.printer_combo.addItems(printers)

        # Set default printer as selected
        default_printer = win32print.GetDefaultPrinter()
        default_index = self.printer_combo.findText(default_printer)
        if default_index >= 0:
            self.printer_combo.setCurrentIndex(default_index)

        # Duplex printing checkbox
        self.duplex_checkbox = QCheckBox("Impressão frente e verso")
        self.duplex_checkbox.setChecked(True)

        # Buttons
        buttons_layout = QVBoxLayout()
        print_button = QPushButton("Imprimir")
        cancel_button = QPushButton("Cancelar")

        print_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)

        # Add widgets to layout
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


class PrintProductionOrderDialogV3(QtWidgets.QDialog):
    def __init__(self, df: pd.DataFrame, df_op_table, selected_row=None, parent=None):
        super().__init__(parent)
        self.label_status = None
        self.pdf_thread = None
        self.progress_bar = None
        self.layout = None
        self.df = df
        self.df_op_table = df_op_table
        self.selected_row = selected_row
        self.generated_pdfs = []
        self.init_ui()
        registrar_fonte_personalizada()
        self.print_production_order()

    def init_ui(self):
        self.setWindowTitle("Eureka® PCP - Impressão de OP")
        self.setGeometry(100, 100, 400, 100)

        self.layout = QVBoxLayout(self)

        self.label_status = QLabel("Publicando Ordem de Produção...", self)

        # Progress bar
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setRange(0, 100)

        self.layout.addWidget(self.label_status)
        self.layout.addWidget(self.progress_bar)

        self.setLayout(self.layout)
        self.center()

    def center(self):
        qr = self.frameGeometry()
        cp = QtWidgets.QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def print_production_order(self):
        output_dir = r"\\192.175.175.4\dados\EMPRESA\PRODUCAO\ORDEM_DE_PRODUCAO"
        os.makedirs(output_dir, exist_ok=True)

        self.pdf_thread = PDFGeneratorThread(self.df, self.df_op_table, output_dir, self.selected_row)
        self.pdf_thread.finished.connect(self.on_pdf_generation_complete)
        self.pdf_thread.error.connect(self.on_pdf_generation_error)
        self.pdf_thread.progress.connect(lambda value, current, total: self.update_progress(value, current, total))
        self.pdf_thread.start()

    def update_progress(self, value, current, total):
        self.progress_bar.setValue(value)
        self.label_status.setText(f"Publicando OP {current} de {total}")

    def print_pdf_with_sumatrapdf(self, printer_name, pdf_path, duplex=True):
        try:
            # Path to SumatraPDF (you might need to adjust this path)
            sumatra_path = r"C:\Users\Eliezer\AppData\Local\SumatraPDF.exe"

            # Build the command
            command = [
                sumatra_path,
                "-print-to", printer_name,
                "-print-settings", f"duplex={duplex and 'yes' or 'no'}",
                pdf_path
            ]

            # Run the command
            process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            stdout, stderr = process.communicate()

            if process.returncode == 0:
                return True
            else:
                raise Exception(f"Error printing: {stderr.decode()}")

        except Exception as e:
            QMessageBox.warning(
                self,
                "Erro na Impressão",
                f"Erro ao imprimir {os.path.basename(pdf_path)}: {str(e)}"
            )
            return False

    def print_pdfs_with_default_program(self, printer_name, pdf_path):
        try:
            # Set the specified printer as default
            previous_printer = win32print.GetDefaultPrinter()
            win32print.SetDefaultPrinter(printer_name)

            # Print using the default Windows print command
            win32api.ShellExecute(
                0,
                "print",
                pdf_path,
                None,
                ".",
                0
            )

            # Restore the previous default printer
            win32print.SetDefaultPrinter(previous_printer)
            return True

        except Exception as e:
            QMessageBox.warning(
                self,
                "Erro na Impressão",
                f"Erro ao imprimir {os.path.basename(pdf_path)}: {str(e)}"
            )
            return False

    def print_pdfs(self):
        if not self.generated_pdfs:
            QMessageBox.warning(self, "Aviso", "Nenhum PDF encontrado para impressão!")
            return

        try:
            printer_dialog = PrinterSelectionDialog(self)
            if printer_dialog.exec_() == QDialog.Accepted:
                selected_printer = printer_dialog.get_selected_printer()
                is_duplex = printer_dialog.is_duplex_selected()

                success_count = 0
                total_files = len(self.generated_pdfs)

                # First try with SumatraPDF
                try:
                    for pdf_path in self.generated_pdfs:
                        if self.print_pdf_with_sumatrapdf(selected_printer, pdf_path, is_duplex):
                            success_count += 1
                except Exception:
                    # If SumatraPDF fails, fall back to default Windows printing
                    for pdf_path in self.generated_pdfs:
                        if self.print_pdfs_with_default_program(selected_printer, pdf_path):
                            success_count += 1

                if success_count > 0:
                    QMessageBox.information(
                        self,
                        "Sucesso",
                        f"{success_count} de {total_files} documentos enviados para impressora {selected_printer}"
                    )
                else:
                    QMessageBox.warning(
                        self,
                        "Aviso",
                        "Nenhum documento foi impresso devido a erros."
                    )

        except Exception as e:
            QMessageBox.critical(
                self,
                "Erro",
                f"Erro ao iniciar impressão: {str(e)}"
            )

    def on_pdf_generation_complete(self):
        # Store the generated PDF paths
        self.generated_pdfs = []
        output_dir = r"\\192.175.175.4\dados\EMPRESA\PRODUCAO\ORDEM_DE_PRODUCAO"
        for file in os.listdir(output_dir):
            if file.endswith('.pdf'):
                self.generated_pdfs.append(os.path.join(output_dir, file))

        # Ask if user wants to print
        reply = QMessageBox.question(
            self,
            'Impressão',
            'Deseja enviar para impressora?',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.print_pdfs()

        information_dialog(
            self,
            "Eureka® PCP - Impressão de OP",
            "Processo finalizado com sucesso! ✅🥳\n\n"
            "Os arquivos foram salvos em:\n"
            r"\192.175.175.4\dados\EMPRESA\PRODUCAO\ORDEM_DE_PRODUCAO"
        )
        self.close()

    # def on_pdf_generation_complete(self):
    #     information_dialog(self, "Eureka® PCP - Impressão de OP", "Processo finalizado com sucesso! ✅🥳\n\n"
    #                                                           "Os arquivos foram salvos em:\n"
    #                                                           r"\192.175.175.4\dados\EMPRESA\PRODUCAO\ORDEM_DE_PRODUCAO")
    #     self.close()

    def on_pdf_generation_error(self, error):
        QMessageBox.critical(self, "Eureka® PCP - Erro", f"Erro ao gerar PDF ❌\nErro: {error}")
        self.close()
