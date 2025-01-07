import sqlite3
from PyQt5.QtWidgets import QLineEdit, QApplication, QMainWindow

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.conn = sqlite3.connect("search_history.db")
        self.create_table()
        self.line_edit = QLineEdit(self)
        self.line_edit.setPlaceholderText("Pesquisar...")
        self.line_edit.textChanged.connect(self.show_history)
        self.setCentralWidget(self.line_edit)

    def create_table(self):
        with self.conn:
            self.conn.execute(
                "CREATE TABLE IF NOT EXISTS search_history (query TEXT UNIQUE)"
            )

    def show_history(self, text):
        if text:
            cursor = self.conn.cursor()
            cursor.execute(
                "SELECT query FROM search_history WHERE query LIKE ? ORDER BY query",
                (f"{text}%",),
            )
            suggestions = [row[0] for row in cursor.fetchall()]
            print("Sugestões:", suggestions)

    def closeEvent(self, event):
        text = self.line_edit.text().strip()
        if text:
            with self.conn:
                self.conn.execute(
                    "INSERT OR IGNORE INTO search_history (query) VALUES (?)", (text,)
                )

app = QApplication([])
window = MainWindow()
window.show()
app.exec_()
