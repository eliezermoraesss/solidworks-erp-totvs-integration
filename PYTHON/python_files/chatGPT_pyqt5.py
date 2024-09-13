import sys
import requests
from PyQt5.QtWidgets import QApplication, QVBoxLayout, QWidget, QPushButton, QTextEdit, QLineEdit

class ChatGPTApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('ChatGPT API Integration')

        self.layout = QVBoxLayout()

        # Campo de entrada para enviar uma mensagem
        self.input_text = QLineEdit(self)
        self.input_text.setPlaceholderText('Digite sua mensagem...')
        self.layout.addWidget(self.input_text)

        # Botão para enviar a mensagem
        self.send_button = QPushButton('Enviar para ChatGPT', self)
        self.send_button.clicked.connect(self.call_chatgpt)
        self.layout.addWidget(self.send_button)

        # Campo de texto para exibir a resposta
        self.output_text = QTextEdit(self)
        self.output_text.setReadOnly(True)
        self.layout.addWidget(self.output_text)

        self.setLayout(self.layout)

    def call_chatgpt(self):
        user_input = self.input_text.text()
        if not user_input:
            return

        # Chamada à API do ChatGPT
        api_key = 'SUA_API_KEY_AQUI'  # Substitua com sua chave de API do OpenAI
        headers = {
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json'
        }
        data = {
            "model": "gpt-3.5-turbo",
            "messages": [{"role": "user", "content": user_input}],
        }

        response = requests.post('https://api.openai.com/v1/chat/completions', headers=headers, json=data)

        if response.status_code == 200:
            result = response.json()
            chatgpt_reply = result['choices'][0]['message']['content']
            self.output_text.setText(chatgpt_reply)
        else:
            self.output_text.setText('Erro ao chamar a API do ChatGPT.')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ChatGPTApp()
    window.show()
    sys.exit(app.exec_())
