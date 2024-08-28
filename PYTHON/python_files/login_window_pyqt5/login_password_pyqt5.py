from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QLineEdit, QLabel, QPushButton, QVBoxLayout, QMessageBox, QComboBox
import bcrypt
import random
import sqlite3
from datetime import datetime, timedelta
import sys
import os


class AuthController:
    def __init__(self):
        # Define o caminho para salvar o banco de dados na pasta database
        database_path = os.path.join(os.path.dirname(__file__), 'database', 'app.db')
        os.makedirs(os.path.dirname(database_path), exist_ok=True)
        self.conn = sqlite3.connect(database_path)
        self.create_tables()

    def create_tables(self):
        cursor = self.conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS users (
                          id INTEGER PRIMARY KEY AUTOINCREMENT,
                          username TEXT NOT NULL UNIQUE,
                          email TEXT NOT NULL UNIQUE,
                          hashed_password TEXT NOT NULL,
                          role TEXT NOT NULL)''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS password_reset (
                          id INTEGER PRIMARY KEY AUTOINCREMENT,
                          user_id INTEGER,
                          reset_code TEXT NOT NULL,
                          expiration_time DATETIME NOT NULL)''')
        self.conn.commit()

    def hash_password(self, password):
        salt = bcrypt.gensalt()
        return bcrypt.hashpw(password.encode('utf-8'), salt)

    def verify_password(self, hashed_password, password):
        return bcrypt.checkpw(password.encode('utf-8'), hashed_password)

    def create_user(self, username, email, password, role):
        if self.get_user_by_username(username):
            return False, "O nome de usuário já está em uso."
        if self.get_user_by_email(email):
            return False, "O e-mail já está em uso."

        hashed_password = self.hash_password(password)
        cursor = self.conn.cursor()
        cursor.execute('''INSERT INTO users (username, email, hashed_password, role)
                          VALUES (?, ?, ?, ?)''', (username, email, hashed_password, role))
        self.conn.commit()
        return True, "Usuário registrado com sucesso!"

    def get_user_by_username(self, username):
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM users WHERE username = ?", (username,))
        return cursor.fetchone()

    def get_user_by_email(self, email):
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM users WHERE email = ?", (email,))
        return cursor.fetchone()

    def generate_reset_code(self, email):
        user = self.get_user_by_email(email)
        if user:
            reset_code = str(random.randint(100000, 999999))
            expiration_time = datetime.now() + timedelta(minutes=2)
            cursor = self.conn.cursor()
            cursor.execute('''INSERT INTO password_reset (user_id, reset_code, expiration_time)
                              VALUES (?, ?, ?)''', (user[0], reset_code, expiration_time))
            self.conn.commit()
            self.send_reset_email(email, reset_code)
            return True
        return False

    def send_reset_email(self, email, reset_code):
        print(f'Send reset code {reset_code} to {email}')

    def verify_reset_code(self, email, code):
        user = self.get_user_by_email(email)
        if not user:
            return False, "Email inválido."

        cursor = self.conn.cursor()
        cursor.execute('''SELECT reset_code, expiration_time FROM password_reset
                          WHERE user_id = ? ORDER BY id DESC LIMIT 1''', (user[0],))
        result = cursor.fetchone()

        if not result:
            return False, "Nenhum código de recuperação encontrado."

        if result[0] != code:
            return False, "Código inválido."

        if datetime.now() > datetime.fromisoformat(result[1]):
            return False, "O código expirou."

        return True, "Código verificado."

    def reset_password(self, email, new_password):
        user = self.get_user_by_email(email)
        if user:
            hashed_password = self.hash_password(new_password)
            cursor = self.conn.cursor()
            cursor.execute('''UPDATE users SET hashed_password = ?
                              WHERE email = ?''', (hashed_password, email))
            self.conn.commit()
            return True
        return False

    def check_authorization(self, username, required_role):
        user = self.get_user_by_username(username)
        if user:
            user_role = user[4]
            if user_role == 'Admin' or user_role == required_role:
                return True
        return False


class LoginWindow(QtWidgets.QWidget):
    def __init__(self, auth_controller):
        super().__init__()
        self.auth_controller = auth_controller
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Login')
        self.setGeometry(100, 100, 300, 200)

        self.username_input = QLineEdit(self)
        self.username_input.setPlaceholderText('Nome de Usuário')

        self.password_input = QLineEdit(self)
        self.password_input.setPlaceholderText('Senha')
        self.password_input.setEchoMode(QLineEdit.Password)

        self.login_button = QPushButton('Login', self)
        self.login_button.clicked.connect(self.login)

        self.forgot_password_button = QPushButton('Esqueceu a senha?', self)
        self.forgot_password_button.clicked.connect(self.forgot_password)

        self.register_button = QPushButton('Registrar', self)
        self.register_button.clicked.connect(self.open_register_window)

        layout = QVBoxLayout()
        layout.addWidget(QLabel('Nome de Usuário:'))
        layout.addWidget(self.username_input)
        layout.addWidget(QLabel('Senha:'))
        layout.addWidget(self.password_input)
        layout.addWidget(self.login_button)
        layout.addWidget(self.forgot_password_button)
        layout.addWidget(self.register_button)

        self.setLayout(layout)

    def login(self):
        username = self.username_input.text()
        password = self.password_input.text()
        user = self.auth_controller.get_user_by_username(username)

        if user and self.auth_controller.verify_password(user[3], password):
            self.main_window = MainWindow(self.auth_controller, username)  # Abre a janela principal se o login for bem-sucedido
            self.main_window.show()
            self.close()
        else:
            QMessageBox.warning(self, 'Erro', 'Credenciais inválidas')

    def forgot_password(self):
        email, ok = QtWidgets.QInputDialog.getText(self, 'Recuperar Senha', 'Digite seu e-mail:')
        if ok and email:
            if self.auth_controller.generate_reset_code(email):
                self.code_verification_window = CodeVerificationWindow(self.auth_controller, email)
                self.code_verification_window.show()
            else:
                QMessageBox.warning(self, 'Erro', 'Email não encontrado')

    def open_register_window(self):
        self.register_window = RegisterWindow(self.auth_controller)
        self.register_window.show()


class RegisterWindow(QtWidgets.QWidget):
    def __init__(self, auth_controller):
        super().__init__()
        self.auth_controller = auth_controller
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Registrar')
        self.setGeometry(100, 100, 300, 350)

        self.username_input = QLineEdit(self)
        self.username_input.setPlaceholderText('Nome de Usuário')

        self.email_input = QLineEdit(self)
        self.email_input.setPlaceholderText('Email')

        self.password_input = QLineEdit(self)
        self.password_input.setPlaceholderText('Senha')
        self.password_input.setEchoMode(QLineEdit.Password)

        self.confirm_password_input = QLineEdit(self)
        self.confirm_password_input.setPlaceholderText('Confirme a Senha')
        self.confirm_password_input.setEchoMode(QLineEdit.Password)

        self.role_combobox = QComboBox(self)
        self.role_combobox.addItems(["Engenharia", "PCP", "Faturamento", "Comercial", "Compras", "RH", "Almoxarifado", "Expedição"])

        self.register_button = QPushButton('Registrar', self)
        self.register_button.clicked.connect(self.register)

        layout = QVBoxLayout()
        layout.addWidget(QLabel('Nome de Usuário:'))
        layout.addWidget(self.username_input)
        layout.addWidget(QLabel('Email:'))
        layout.addWidget(self.email_input)
        layout.addWidget(QLabel('Senha:'))
        layout.addWidget(self.password_input)
        layout.addWidget(QLabel('Confirme a Senha:'))
        layout.addWidget(self.confirm_password_input)
        layout.addWidget(QLabel('Setor:'))
        layout.addWidget(self.role_combobox)
        layout.addWidget(self.register_button)

        self.setLayout(layout)

    def register(self):
        username = self.username_input.text()
        email = self.email_input.text()
        password = self.password_input.text()
        confirm_password = self.confirm_password_input.text()
        role = self.role_combobox.currentText()

        if password != confirm_password:
            QMessageBox.warning(self, 'Erro', 'As senhas não coincidem')
            return

        success, message = self.auth_controller.create_user(username, email, password, role)
        if success:
            QMessageBox.information(self, 'Sucesso', message)
            self.close()
        else:
            QMessageBox.warning(self, 'Erro', message)


class CodeVerificationWindow(QtWidgets.QWidget):
    def __init__(self, auth_controller, email):
        super().__init__()
        self.auth_controller = auth_controller
        self.email = email
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Verificação de Código')
        self.setGeometry(100, 100, 300, 200)

        self.code_input = QLineEdit(self)
        self.code_input.setPlaceholderText('Código de verificação')

        self.verify_button = QPushButton('Verificar', self)
        self.verify_button.clicked.connect(self.verify_code)

        layout = QVBoxLayout()
        layout.addWidget(QLabel('Digite o código enviado para o seu e-mail:'))
        layout.addWidget(self.code_input)
        layout.addWidget(self.verify_button)
        self.setLayout(layout)

    def verify_code(self):
        code = self.code_input.text()
        success, message = self.auth_controller.verify_reset_code(self.email, code)
        if success:
            self.reset_password_window = ResetPasswordWindow(self.auth_controller, self.email)
            self.reset_password_window.show()
            self.close()
        else:
            QMessageBox.warning(self, 'Erro', message)


class ResetPasswordWindow(QtWidgets.QWidget):
    def __init__(self, auth_controller, email):
        super().__init__()
        self.auth_controller = auth_controller
        self.email = email
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Redefinir Senha')
        self.setGeometry(100, 100, 300, 250)

        self.new_password_input = QLineEdit(self)
        self.new_password_input.setPlaceholderText('Nova senha')
        self.new_password_input.setEchoMode(QLineEdit.Password)

        self.confirm_password_input = QLineEdit(self)
        self.confirm_password_input.setPlaceholderText('Confirme a nova senha')
        self.confirm_password_input.setEchoMode(QLineEdit.Password)

        self.confirm_button = QPushButton('Redefinir Senha', self)
        self.confirm_button.clicked.connect(self.reset_password)

        layout = QVBoxLayout()
        layout.addWidget(QLabel('Digite a nova senha:'))
        layout.addWidget(self.new_password_input)
        layout.addWidget(QLabel('Confirme a nova senha:'))
        layout.addWidget(self.confirm_password_input)
        layout.addWidget(self.confirm_button)
        self.setLayout(layout)

    def reset_password(self):
        new_password = self.new_password_input.text()
        confirm_password = self.confirm_password_input.text()

        if new_password != confirm_password:
            QMessageBox.warning(self, 'Erro', 'As senhas não coincidem')
            return

        if self.auth_controller.reset_password(self.email, new_password):
            QMessageBox.information(self, 'Sucesso', 'Senha redefinida com sucesso!')
            self.close()
        else:
            QMessageBox.warning(self, 'Erro', 'Erro ao redefinir a senha.')


class MainWindow(QtWidgets.QWidget):
    def __init__(self, auth_controller, username):
        super().__init__()
        self.auth_controller = auth_controller
        self.username = username
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Main Window')
        self.setGeometry(100, 100, 400, 300)

        # Exemplo de como usar a verificação de autorização
        if self.auth_controller.check_authorization(self.username, 'Engenharia'):
            QLabel("Acesso concedido à função específica de Engenharia.", self).show()
        else:
            QLabel("Acesso negado.", self).show()

        layout = QVBoxLayout()
        layout.addWidget(QLabel(f"Bem-vindo, {self.username}!"))
        self.setLayout(layout)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    auth_controller = AuthController()
    login_window = LoginWindow(auth_controller)
    login_window.show()
    sys.exit(app.exec_())
