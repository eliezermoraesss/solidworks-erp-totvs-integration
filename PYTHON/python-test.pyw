import ctypes

# Exibe uma caixa de diálogo com a mensagem "Hello, World!"
ctypes.windll.user32.MessageBoxW(0, "Olá do seu amigo Python!", "Janelinha da alegria", 1)
