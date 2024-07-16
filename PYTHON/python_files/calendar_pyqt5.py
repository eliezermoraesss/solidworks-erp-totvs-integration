from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QDateEdit
from PyQt5.QtCore import QDate
import sys

class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        
        # Configurações da janela
        self.setWindowTitle("Exemplo de QDateEdit")
        self.setGeometry(100, 100, 300, 200)
        
        # Layout
        layout = QVBoxLayout()
        
        # Campo de data
        self.campo_data_inicio = QDateEdit(self)
        self.campo_data_inicio.setCalendarPopup(True)  # Ativa o popup do calendário
        data_atual = QDate.currentDate()
        meses_a_remover = 3
        data_inicio = data_atual.addMonths(-meses_a_remover)
        self.campo_data_inicio.setDate(data_inicio)
        
        # Adiciona ao layout
        layout.addWidget(self.campo_data_inicio)
        
        # Define o layout da janela
        self.setLayout(layout)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")  # Tenta um estilo diferente
    window = MyApp()
    window.show()
    sys.exit(app.exec_())
