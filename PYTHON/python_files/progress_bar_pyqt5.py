import sys
import time
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QProgressBar, QPushButton
from PyQt5.QtCore import Qt, QThread, pyqtSignal

class Worker(QThread):
    progress = pyqtSignal(int)

    def run(self):
        for i in range(101):
            time.sleep(0.1)  # Simula uma tarefa demorada
            self.progress.emit(i)

class Window(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setGeometry(100, 100, 400, 200)
        self.setWindowTitle('Barra de Progresso')

        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        self.progressBar = QProgressBar(self)
        self.progressBar.setMaximum(100)
        self.layout.addWidget(self.progressBar)

        self.startButton = QPushButton('Iniciar', self)
        self.startButton.clicked.connect(self.startTask)
        self.layout.addWidget(self.startButton)

    def startTask(self):
        self.thread = Worker()
        self.thread.progress.connect(self.updateProgress)
        self.thread.start()

    def updateProgress(self, value):
        self.progressBar.setValue(value)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec_())
