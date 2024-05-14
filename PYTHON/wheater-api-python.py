import requests

class ConsultaApp:
    def __init__(self):
        # Replace 'YOUR_API_KEY' with your actual OpenWeatherMap API key
        self.api_key = '2a52e6f30c7001670428fb9b7a253c36'

    def obter_tempo_atual(self, cidade):
        url = f"http://api.openweathermap.org/data/2.5/weather?q={cidade}&appid={self.api_key}"
        response = requests.get(url)
        
        if response.status_code == 200:
            data = response.json()
            temperatura = data['main']['temp']
            descricao = data['weather'][0]['description']
            clima = f"{descricao.capitalize()}, {temperatura:.1f}°C"
            print(data)
            print(type(clima))
            return clima
        else:
            return "Erro ao obter informações de clima"

    def imprimir_tempo_atual(self):
        # Replace 'CITY_NAME' with your desired city
        weather_info = self.obter_tempo_atual('Mogi Mirim, BR')
        print(weather_info)

if __name__ == "__main__":
    # Replace 'YOUR_API_KEY' with your actual OpenWeatherMap API key
    api_key = '17afec878c0acf4acb5ab176fb513d60'
    
    # Create an instance of ConsultaApp
    consulta_app = ConsultaApp()

    # Call the method to print weather information
    consulta_app.imprimir_tempo_atual()
