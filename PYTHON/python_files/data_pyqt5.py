from PyQt5.QtCore import QDate, QDateTime
from PyQt5.QtWidgets import QApplication, QWidget, QLineEdit, QCalendarWidget, QVBoxLayout, QPushButton

class DataWidget(QWidget):
    def __init__(self):
        super().__init__()

        # Criar layout vertical
        self.layout = QVBoxLayout()

        # Criar campo de data
        self.data_lineEdit = QLineEdit()
        self.data_lineEdit.setPlaceholderText("Data")

        # Criar botão "Hoje"
        self.hoje_button = QPushButton("Hoje")
        self.hoje_button.clicked.connect(self.set_hoje)

        # Criar calendário (oculto inicialmente)
        self.calendario = QCalendarWidget()
        self.calendario.setVisible(False)

        # Conectar sinal de data selecionada do calendário
        self.calendario.selectionChanged.connect(self.set_data_lineEdit)

        # Adicionar widgets ao layout
        self.layout.addWidget(self.data_lineEdit)
        self.layout.addWidget(self.hoje_button)
        self.layout.addWidget(self.calendario)

        # Definir layout como layout principal
        self.setLayout(self.layout)

    def set_hoje(self):
        """Define a data de hoje no campo de linha."""
        data_hoje = QDate.currentDate()
        self.data_lineEdit.setText(data_hoje.toString("dd/MM/yyyy"))

    def set_data_lineEdit(self):
        """Define a data selecionada no calendário no campo de linha."""
        data_selecionada = self.calendario.selectedDate()
        self.data_lineEdit.setText(data_selecionada.toString("dd/MM/yyyy"))

    def show_calendar(self):
        """Mostrar o calendário."""
        self.calendario.setVisible(True)

    def hide_calendar(self):
        """Ocultar o calendário."""
        self.calendario.setVisible(False)
if __name__ == "__main__":
    app = QApplication([])
    widget = DataWidget()
    widget.show()
    app.exec()
