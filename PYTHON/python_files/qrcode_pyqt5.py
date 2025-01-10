import sys
from io import BytesIO

from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QVBoxLayout, QWidget, QPushButton, QLineEdit, QFileDialog
from PyQt5.QtGui import QPixmap
import qrcode


class QRCodeApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerador de QR Code")

        self.layout = QVBoxLayout()

        self.input_label = QLabel("Digite o conteúdo:")
        self.layout.addWidget(self.input_label)

        self.input_field = QLineEdit()
        self.layout.addWidget(self.input_field)

        self.generate_button = QPushButton("Gerar QR Code")
        self.generate_button.clicked.connect(self.generate_qr_code)
        self.layout.addWidget(self.generate_button)

        self.qr_label = QLabel()
        self.layout.addWidget(self.qr_label)

        self.save_button = QPushButton("Salvar QR Code")
        self.save_button.clicked.connect(self.save_qr_code)
        self.layout.addWidget(self.save_button)

        container = QWidget()
        container.setLayout(self.layout)
        self.setCentralWidget(container)

    def generate_qr_code(self):
        content = self.input_field.text()
        if content:
            qr_img = qrcode.make(content)
            # Converta para QPixmap diretamente
            buffer = BytesIO()
            qr_img.save(buffer, format="PNG")
            pixmap = QPixmap()
            pixmap.loadFromData(buffer.getvalue())
            self.qr_label.setPixmap(pixmap)

    def save_qr_code(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self, "Salvar QR Code", "", "PNG Files (*.png)", options=options)
        if file_path:
            content = self.input_field.text()
            qr_img = qrcode.make(content)
            qr_img.save(file_path)


app = QApplication(sys.argv)
window = QRCodeApp()
window.show()
sys.exit(app.exec_())
