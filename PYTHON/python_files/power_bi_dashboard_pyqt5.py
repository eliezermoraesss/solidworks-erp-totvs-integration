import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import QUrl


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Power BI Viewer")

        # Cria o widget QWebEngineView
        self.browser = QWebEngineView()

        # Converte a string da URL para um objeto QUrl
        power_bi_url = QUrl("https://app.powerbi.com/groups/me/reports/fc538d79-9f1d-4798-af08-bbc764b11b4d?ctid=1a30606f-47bf-4606-aa81-7245533ad2d9&pbi_source=linkShare&bookmarkGuid=82e16595-3a31-40b8-bf12-85869a6a8af1")

        # Carrega a URL do relatório do Power BI
        self.browser.setUrl(power_bi_url)

        # Layout para o QWebEngineView
        layout = QVBoxLayout()
        layout.addWidget(self.browser)

        # Define o layout no widget central
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
