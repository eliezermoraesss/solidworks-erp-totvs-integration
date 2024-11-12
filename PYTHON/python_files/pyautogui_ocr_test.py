import cv2
import pyautogui
import numpy as np
import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk

class ImageComparatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Comparador de Imagens")
        
        self.sensitivity = tk.IntVar(value=30)
        self.image1_path = ""
        self.image2_path = ""
        
        # Botões e controle deslizante
        tk.Button(root, text="Capturar Imagem 1", command=self.capture_image1).pack(pady=5)
        tk.Button(root, text="Capturar Imagem 2", command=self.capture_image2).pack(pady=5)
        
        tk.Label(root, text="Ajuste a Sensibilidade:").pack(pady=5)
        tk.Scale(root, from_=0, to=100, orient="horizontal", variable=self.sensitivity).pack(pady=5)
        
        tk.Button(root, text="Comparar Imagens", command=self.compare_images).pack(pady=5)
        self.image_label = tk.Label(root)
        self.image_label.pack()

    def capture_image1(self):
        self.image1_path = "image1.png"
        self.capture_screenshot(self.image1_path)
        print("Imagem 1 capturada.")

    def capture_image2(self):
        self.image2_path = "image2.png"
        self.capture_screenshot(self.image2_path)
        print("Imagem 2 capturada.")

    def capture_screenshot(self, filename):
        screenshot = pyautogui.screenshot()
        screenshot.save(filename)

    def compare_images(self):
        if not self.image1_path or not self.image2_path:
            print("Capture ambas as imagens antes de comparar.")
            return
        
        threshold_value = self.sensitivity.get()
        output_path = "output_diff.png"
        self.find_differences(self.image1_path, self.image2_path, output_path, threshold_value)
        
        # Carregar e exibir a imagem de saída com as diferenças
        img = Image.open(output_path)
        img_tk = ImageTk.PhotoImage(img)
        self.image_label.config(image=img_tk)
        self.image_label.image = img_tk

    def preprocess_image(self, image_path):
        image = cv2.imread(image_path)
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        gray = cv2.GaussianBlur(gray, (5, 5), 0)
        return gray

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

# Inicializa a aplicação
root = tk.Tk()
app = ImageComparatorApp(root)
root.mainloop()
