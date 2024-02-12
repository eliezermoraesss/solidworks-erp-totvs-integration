import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton
from PyQt5.QtCore import Qt

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle("Dark Theme Example")
        self.setGeometry(100, 100, 400, 300)

        # Botão com estilo personalizado
        btn = QPushButton("Botão", self)
        btn.setGeometry(150, 150, 100, 50)
        btn.setStyleSheet("""
            QPushButton {
                background-color: #333;
                color: #fff;
                border-radius: 10px;
            }
            QPushButton:hover {
                background-color: #555;
            }
            QPushButton:pressed {
                background-color: #777;
            }
        """)

if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Aplicando um estilo escuro globalmente para toda a aplicação
    app.setStyle('Fusion')
    dark_palette = app.palette()
    dark_palette.setColor(app.palette().Window, Qt.black)
    dark_palette.setColor(app.palette().WindowText, Qt.white)
    app.setPalette(dark_palette)

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())
