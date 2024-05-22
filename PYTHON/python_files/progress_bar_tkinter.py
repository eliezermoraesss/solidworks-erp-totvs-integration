import tkinter as tk
from tkinter import ttk
import threading
import time

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Barra de Progresso")

        self.progress = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=20)

        self.start_button = tk.Button(root, text="Iniciar", command=self.start_task)
        self.start_button.pack(pady=20)

    def start_task(self):
        self.thread = threading.Thread(target=self.run_task)
        self.thread.start()

    def run_task(self):
        for i in range(101):
            time.sleep(0.1)  # Simula uma tarefa demorada
            self.update_progress(i)

    def update_progress(self, value):
        self.progress['value'] = value
        self.root.update_idletasks()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.geometry("400x200")
    root.mainloop()
