import os
import sqlite3
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QLineEdit,
    QVBoxLayout, QWidget, QCompleter
)
from PyQt5.QtCore import Qt, QStringListModel


class SearchHistoryManager:
    def __init__(self, db_path='search_history.db'):
        app_data = os.getenv('LOCALAPPDATA') or os.path.expanduser('~\\AppData\\Roaming')
        # Determina o caminho do banco de dados
        self.db_path = os.path.join(app_data, 'Eureka', db_path)

        # Cria diretório se não existir
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)

        # Conecta ao banco de dados
        self.conn = sqlite3.connect(self.db_path)
        self.cursor = self.conn.cursor()

        # Cria tabela de histórico se não existir
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS search_history (
                field_name TEXT,
                value TEXT,
                timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(field_name, value)
            )
        ''')
        self.conn.commit()

    def save_history(self, field_name, value):
        """Salva item no histórico"""
        if not value:  # Ignora valores vazios
            return

        try:
            # Remove duplicatas e insere novo item
            self.cursor.execute('''
                INSERT OR REPLACE INTO search_history 
                (field_name, value) VALUES (?, ?)
            ''', (field_name, value))
            self.conn.commit()
        except sqlite3.Error as e:
            print(f"Erro ao salvar histórico: {e}")

    def get_history(self, field_name):
        """Recupera histórico de um campo"""
        self.cursor.execute('''
            SELECT value FROM search_history 
            WHERE field_name = ? 
            ORDER BY timestamp DESC
        ''', (field_name))
        return [row[0] for row in self.cursor.fetchall()]

    def clear_history(self, field_name=None):
        """Limpa histórico"""
        if field_name:
            self.cursor.execute(
                'DELETE FROM search_history WHERE field_name = ?',
                (field_name,)
            )
        else:
            self.cursor.execute('DELETE FROM search_history')
        self.conn.commit()

    def __del__(self):
        """Fecha conexão com banco ao destruir objeto"""
        self.conn.close()


class HistorySearchApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Histórico de Pesquisa")
        self.setGeometry(300, 300, 400, 200)

        self.history_manager = SearchHistoryManager()

        container = QWidget()
        layout = QVBoxLayout()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # Dicionário para guardar referência aos campos e completers
        self.fields = {}

        for label, field_name in [
            ("Código", "codigo"),
            ("E-mail", "email"),
            ("Produto", "product")
        ]:
            line_edit = QLineEdit()
            line_edit.setPlaceholderText(label)

            # Cria completer
            completer = QCompleter()
            completer.setCaseSensitivity(Qt.CaseInsensitive)
            completer.setModel(QStringListModel())
            line_edit.setCompleter(completer)

            # Atualiza completer com dados históricos
            self.update_completer(field_name, completer)

            # Guarda referências
            self.fields[field_name] = {
                'line_edit': line_edit,
                'completer': completer
            }

            # Conecta sinal
            line_edit.returnPressed.connect(
                lambda fn=field_name: self.save_search_history(fn)
            )

            layout.addWidget(line_edit)

    def update_completer(self, field_name, completer):
        """Atualiza a lista de sugestões do completer"""
        history = self.history_manager.get_history(field_name)
        completer.model().setStringList(history)

    def save_search_history(self, field_name):
        """Salva histórico e atualiza completer"""
        value = self.fields[field_name]['line_edit'].text()

        if value.strip():  # Verifica se não está vazio
            self.history_manager.save_history(field_name, value)

            # Atualiza completer do campo
            completer = self.fields[field_name]['completer']
            self.update_completer(field_name, completer)


def main():
    app = QApplication([])
    window = HistorySearchApp()
    window.show()
    app.exec_()


if __name__ == "__main__":
    main()
