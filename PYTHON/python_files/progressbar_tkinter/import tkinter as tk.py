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

        self.status_label = tk.Label(root, text="")
        self.status_label.pack(pady=20)

    def start_task(self):
        self.thread = threading.Thread(target=self.run_task)
        self.thread.start()

    def run_task(self):
        total_steps = 5  # Número total de passos/tarefas no script
        for step in range(total_steps):
            self.execute_step(step)
            progress_value = int((step + 1) / total_steps * 100)
            self.update_progress(progress_value)
    
    def execute_step(self, step):
        # Simulando diferentes etapas do script
        if step == 0:
            self.status_label.config(text="Carregando dados...")
            time.sleep(2)  # Simula a execução do passo 1
        elif step == 1:
            self.status_label.config(text="Processando dados...")
            time.sleep(3)  # Simula a execução do passo 2
        elif step == 2:
            self.status_label.config(text="Analisando resultados...")
            time.sleep(1)  # Simula a execução do passo 3
        elif step == 3:
            self.status_label.config(text="Gerando relatório...")
            time.sleep(2)  # Simula a execução do passo 4
        elif step == 4:
            self.status_label.config(text="Concluído!")
            time.sleep(1)  # Simula a execução do passo 5

    def update_progress(self, value):
        self.progress['value'] = value
        self.root.update_idletasks()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.geometry("400x200")
    root.mainloop()
