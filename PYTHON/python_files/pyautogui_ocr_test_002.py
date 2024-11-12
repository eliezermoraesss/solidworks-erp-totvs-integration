import os
import cv2
import pyautogui
import numpy as np
import tkinter as tk
from tkinter import Toplevel
from PIL import Image, ImageTk
from datetime import datetime

class ImageComparatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Controles de Comparação de Imagens")
        
        # Define o caminho padrão para salvar as imagens
        self.image_folder = os.path.join(os.path.expanduser("~"), "Pictures", "ImageCompare")
        os.makedirs(self.image_folder, exist_ok=True)  # Cria a pasta se não existir

        self.sensitivity = tk.IntVar(value=0)
        
        # Botões para capturar e visualizar imagens
        tk.Button(root, text="Capturar Imagem 1", command=self.capture_image1).pack(pady=5)
        tk.Button(root, text="Capturar Imagem 2", command=self.capture_image2).pack(pady=5)
        
        tk.Label(root, text="Ajuste a Sensibilidade:").pack(pady=5)
        tk.Scale(root, from_=0, to=100, orient="horizontal", variable=self.sensitivity).pack(pady=5)
        
        tk.Button(root, text="Comparar Imagens", command=self.compare_images).pack(pady=5)
        
        # Botões para abrir as imagens em janelas separadas
        tk.Button(root, text="Ver Captura 1", command=self.show_image1).pack(pady=5)
        tk.Button(root, text="Ver Captura 2", command=self.show_image2).pack(pady=5)
        tk.Button(root, text="Ver Diferença", command=self.show_difference).pack(pady=5)

        # Armazena os caminhos das imagens
        self.image1_path = ""
        self.image2_path = ""
        self.diff_image_path = ""

    def generate_filename(self, base_name):
        # Gera o nome do arquivo com data e hora atuais
        timestamp = datetime.now().strftime("%d-%m-%Y_%H%M%S")
        return os.path.join(self.image_folder, f"{base_name}_{timestamp}.png")

    def capture_image1(self):
        self.image1_path = self.generate_filename("image1")
        self.capture_screenshot(self.image1_path)
        print("Imagem 1 capturada em:", self.image1_path)

    def capture_image2(self):
        self.image2_path = self.generate_filename("image2")
        self.capture_screenshot(self.image2_path)
        print("Imagem 2 capturada em:", self.image2_path)

    def capture_screenshot(self, filename):
        screenshot = pyautogui.screenshot()
        screenshot.save(filename)

    def preprocess_image(self, image_path):
        image = cv2.imread(image_path)
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        gray = cv2.GaussianBlur(gray, (5, 5), 0)
        return gray

    def compare_images(self):
        if not self.image1_path or not self.image2_path:
            print("Capture ambas as imagens antes de comparar.")
            return
        
        threshold_value = self.sensitivity.get()
        self.diff_image_path = self.generate_filename("output_diff")
        self.find_differences(self.image1_path, self.image2_path, self.diff_image_path, threshold_value)
        print("Diferença salva em:", self.diff_image_path)

    def find_differences(self, image1_path, image2_path, output_path, threshold_value):
        gray1 = self.preprocess_image(image1_path)
        gray2 = self.preprocess_image(image2_path)
        
        diff = cv2.absdiff(gray1, gray2)
        _, thresh = cv2.threshold(diff, threshold_value, 255, cv2.THRESH_BINARY)
        
        contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        image1 = cv2.imread(image1_path)
        for contour in contours:
            if cv2.contourArea(contour) > 10:
                x, y, w, h = cv2.boundingRect(contour)
                cv2.rectangle(image1, (x, y), (x+w, y+h), (0, 255, 0), 2)
        
        cv2.imwrite(output_path, image1)

    def show_image1(self):
        if os.path.exists(self.image1_path):
            self.open_image_window(self.image1_path, "Captura 1")
        else:
            print("Imagem 1 ainda não foi capturada.")

    def show_image2(self):
        if os.path.exists(self.image2_path):
            self.open_image_window(self.image2_path, "Captura 2")
        else:
            print("Imagem 2 ainda não foi capturada.")

    def show_difference(self):
        if os.path.exists(self.diff_image_path):
            self.open_image_window(self.diff_image_path, "Diferença")
        else:
            print("Diferença ainda não foi calculada.")

    def open_image_window(self, image_path, title):
        # Cria uma nova janela para exibir a imagem
        new_window = Toplevel(self.root)
        new_window.title(title)
        
        # Carrega e exibe a imagem na nova janela
        img = Image.open(image_path)
        img_tk = ImageTk.PhotoImage(img)
        
        label = tk.Label(new_window, image=img_tk)
        label.image = img_tk  # Necessário para manter a imagem na memória
        label.pack()

# Inicializa a aplicação
root = tk.Tk()
app = ImageComparatorApp(root)
root.mainloop()
