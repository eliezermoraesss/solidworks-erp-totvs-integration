import os
import tempfile
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import win32print
import win32api

def generate_pdf(output_path):
    c = canvas.Canvas(output_path, pagesize=A4)
    c.drawString(100, 750, "Hello, World!")
    c.save()

def print_pdf(pdf_path):
    printer_name = win32print.GetDefaultPrinter()
    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        hJob = win32print.StartDocPrinter(hPrinter, 1, ("PDF Print Job", None, "RAW"))
        win32print.StartPagePrinter(hPrinter)
        with open(pdf_path, "rb") as f:
            pdf_data = f.read()
            win32print.WritePrinter(hPrinter, pdf_data)
        win32print.EndPagePrinter(hPrinter)
        win32print.EndDocPrinter(hPrinter)
    finally:
        win32print.ClosePrinter(hPrinter)

if __name__ == "__main__":
    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    generate_pdf(temp_pdf.name)
    print_pdf(temp_pdf.name)
    os.remove(temp_pdf.name)
