from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.fonts import addMapping

# Registra todas as variações da fonte Courier
pdfmetrics.registerFont(TTFont('CourierWindows', r'C:\WINDOWS\FONTS\COUR.TTF'))
pdfmetrics.registerFont(TTFont('CourierWindows-Bold', r'C:\WINDOWS\FONTS\COURBD.TTF'))
pdfmetrics.registerFont(TTFont('CourierWindows-Italic', r'C:\WINDOWS\FONTS\COURI.TTF'))
pdfmetrics.registerFont(TTFont('CourierWindows-BoldItalic', r'C:\WINDOWS\FONTS\COURBI.TTF'))

# Mapeia as variações da fonte
addMapping('CourierWindows', 0, 0, 'CourierWindows')           # normal1
addMapping('CourierWindows', 1, 0, 'CourierWindows-Bold')      # bold
addMapping('CourierWindows', 0, 1, 'CourierWindows-Italic')    # italic
addMapping('CourierWindows', 1, 1, 'CourierWindows-BoldItalic')# bold & italic

def criar_pdf_com_courier():
    c = canvas.Canvas("exemplo_courier_completo.pdf")
    
    # Exemplo com todas as variações
    y = 800  # posição vertical inicial
    
    # Normal
    c.setFont('CourierWindows', 12)
    c.drawString(100, y, "Texto Normal com Courier")
    
    # Bold
    c.setFont('CourierWindows-Bold', 12)
    c.drawString(100, y-30, "Texto em Negrito com Courier")
    
    # Italic
    c.setFont('CourierWindows-Italic', 12)
    c.drawString(100, y-60, "Texto em Itálico com Courier")
    
    # Bold Italic
    c.setFont('CourierWindows-BoldItalic', 12)
    c.drawString(100, y-90, "Texto em Negrito e Itálico com Courier")
    
    # Teste com diferentes tamanhos
    c.setFont('CourierWindows', 16)
    c.drawString(100, y-120, "Texto maior (16pt)")
    
    c.setFont('CourierWindows-Bold', 20)
    c.drawString(100, y-150, "Texto grande em negrito (20pt)")
    
    c.save()

# Executa a função
if __name__ == '__main__':
    criar_pdf_com_courier()