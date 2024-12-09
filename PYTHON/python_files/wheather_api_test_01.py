import sys
import requests
from PyQt5.QtWidgets import QApplication, QLabel, QVBoxLayout, QWidget
from PyQt5.QtGui import QPixmap
from io import BytesIO

class WeatherApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Clima em Mogi Mirim")
        self.layout = QVBoxLayout()

        # Labels para exibir texto e ícone
        self.weather_label = QLabel("Buscando clima...")
        self.weather_icon = QLabel()

        # Adicionando labels ao layout
        self.layout.addWidget(self.weather_label)
        self.layout.addWidget(self.weather_icon)

        self.setLayout(self.layout)

        # Chama a função para atualizar o clima
        self.update_weather()

    def update_weather(self):
        api_key = "2a52e6f30c7001670428fb9b7a253c36"
        city = "Mogi Mirim,BR"
        url = f"https://api.openweathermap.org/data/2.5/weather?q={city}&units=metric&lang=pt&appid={api_key}"

        try:
            response = requests.get(url)
            data = response.json()

            if response.status_code == 200:
                # Pega os dados de descrição e temperatura
                description = data["weather"][0]["description"].capitalize()
                temp = data["main"]["temp"]
                icon_code = data["weather"][0]["icon"]

                # Atualiza o texto
                self.weather_label.setText(f"Clima: {description}\nTemperatura: {temp}°C")

                # Baixa e exibe o ícone
                icon_url = f"http://openweathermap.org/img/wn/{icon_code}@2x.png"
                icon_data = requests.get(icon_url).content
                pixmap = QPixmap()
                pixmap.loadFromData(BytesIO(icon_data).read())
                self.weather_icon.setPixmap(pixmap)
            else:
                self.weather_label.setText("Erro ao buscar dados do clima.")

        except Exception as e:
            self.weather_label.setText(f"Erro: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WeatherApp()
    window.show()
    sys.exit(app.exec_())
