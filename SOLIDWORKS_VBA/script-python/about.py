import tkinter as tk
from tkinter import ttk
import time
from tkinter import font as tkfont
import webbrowser

class CadastrarBomTOTVS:
    def __init__(self, window):
        # Configuração inicial da janela
        self.window = window
        window.title("Monitor de progresso")
        window.configure(bg='#111827')  # Equivalente ao bg-gray-900
        window.geometry("400x350")

        # Configuração de estilos
        style = ttk.Style()
        style.theme_use('default')

        # Estilo para a barra de progresso
        style.configure("Custom.Horizontal.TProgressbar",
                        troughcolor='#374151',  # bg-gray-700
                        background='#3B82F6',   # bg-blue-500
                        thickness=15)

        # Frame principal
        main_frame = tk.Frame(window, bg='#111827', padx=20, pady=20)
        main_frame.pack(expand=True, fill='both')

        # Barra de progresso
        self.progress = ttk.Progressbar(
            main_frame,
            orient="horizontal",
            length=300,
            mode="determinate",
            style="Custom.Horizontal.TProgressbar"
        )
        self.progress.pack(pady=20)

        # Labels de status
        self.status_label = tk.Label(
            main_frame,
            text="Cadastrando nova estrutura...",
            bg='#111827',
            fg='white',
            font=('Arial', 10)
        )
        self.status_label.pack(pady=5)

        self.time_label = tk.Label(
            main_frame,
            text="0.000 segundos",
            bg='#111827',
            fg='white',
            font=('Arial', 9)
        )
        self.time_label.pack(pady=2)

        self.brand_label = tk.Label(
            main_frame,
            text="EUREKA®",
            bg='#111827',
            fg='white',
            font=('Arial', 9)
        )
        self.brand_label.pack(pady=2)

        # Botão Sobre
        self.about_button = tk.Button(
            main_frame,
            text="ℹ️ Sobre",
            command=self.show_about,
            bg='#374151',  # bg-gray-700
            fg='white',
            relief='flat',
            padx=15,
            pady=8,
            font=('Arial', 9),
            cursor='hand2'
        )
        self.about_button.pack(pady=20)

        # Inicialização de variáveis
        self.start_time = time.time()
        self.titulo_janela = "CADASTRO DE ESTRUTURA TOTVS®"

    def show_about(self):
        about_window = tk.Toplevel(self.window)
        about_window.title("Sobre - Cadastro BOM TOTVS")
        about_window.geometry("500x400")
        about_window.configure(bg='#111827')

        # Torna a janela modal
        about_window.transient(self.window)
        about_window.grab_set()

        # Frame principal do about
        about_frame = tk.Frame(about_window, bg='#111827', padx=30, pady=20)
        about_frame.pack(expand=True, fill='both')

        # Título
        title_label = tk.Label(
            about_frame,
            text="Cadastro BOM TOTVS - Informações",
            font=('Arial', 14, 'bold'),
            bg='#111827',
            fg='white'
        )
        title_label.pack(pady=(0, 20))

        # Frame para informações em duas colunas
        info_frame = tk.Frame(about_frame, bg='#111827')
        info_frame.pack(fill='x', pady=10)

        # Coluna da esquerda - Contato
        left_frame = tk.Frame(info_frame, bg='#111827')
        left_frame.pack(side='left', padx=10)

        tk.Label(
            left_frame,
            text="Contato",
            font=('Arial', 10, 'bold'),
            bg='#111827',
            fg='white'
        ).pack(anchor='w')

        tk.Label(
            left_frame,
            text="Email: seu.email@empresa.com",
            bg='#111827',
            fg='white',
            font=('Arial', 9)
        ).pack(anchor='w')

        tk.Label(
            left_frame,
            text="Teams: Seu Nome",
            bg='#111827',
            fg='white',
            font=('Arial', 9)
        ).pack(anchor='w')

        # Coluna da direita - Técnico
        right_frame = tk.Frame(info_frame, bg='#111827')
        right_frame.pack(side='right', padx=10)

        tk.Label(
            right_frame,
            text="Técnico",
            font=('Arial', 10, 'bold'),
            bg='#111827',
            fg='white'
        ).pack(anchor='w')

        tk.Label(
            right_frame,
            text="Versão: 1.0.0",
            bg='#111827',
            fg='white',
            font=('Arial', 9)
        ).pack(anchor='w')

        tk.Label(
            right_frame,
            text="Licença: MIT",
            bg='#111827',
            fg='white',
            font=('Arial', 9)
        ).pack(anchor='w')

        # Seção de agradecimentos
        tk.Label(
            about_frame,
            text="\nUm grande obrigado para",
            font=('Arial', 10, 'bold'),
            bg='#111827',
            fg='white'
        ).pack(pady=(20, 10))

        thanks_frame = tk.Frame(about_frame, bg='#111827')
        thanks_frame.pack(fill='x', padx=30)

        for item in ["Time de Engenharia", "Time de TI", "Todos os usuários que forneceram feedback"]:
            tk.Label(
                thanks_frame,
                text=f"• {item}",
                bg='#111827',
                fg='white',
                font=('Arial', 9)
            ).pack(anchor='w')

        # Rodapé
        tk.Label(
            about_frame,
            text="\n© 2024 Sua Empresa",
            bg='#111827',
            fg='white',
            font=('Arial', 9)
        ).pack()

        tk.Label(
            about_frame,
            text="EUREKA®",
            bg='#111827',
            fg='white',
            font=('Arial', 9)
        ).pack()

    def update_progress(self, value):
        self.progress['value'] = value
        self.window.update_idletasks()

        # Atualiza o tempo decorrido
        current_time = time.time()
        elapsed = current_time - self.start_time
        self.time_label.config(text=f"{elapsed:.3f} segundos")

def main():
    root = tk.Tk()
    app = CadastrarBomTOTVS(root)

    # Centraliza a janela na tela
    window_width = 400
    window_height = 350
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")

    root.mainloop()

if __name__ == "__main__":
    main()