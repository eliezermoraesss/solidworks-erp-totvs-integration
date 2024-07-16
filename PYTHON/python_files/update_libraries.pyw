import subprocess
import sys
from pathlib import Path
import tkinter as tk
from tkinter import ttk
import importlib.metadata
import chardet


def install_package(package, text_widget):
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        text_widget.insert(tk.END, f"{package} instalado com sucesso.\n")
    except subprocess.CalledProcessError as e:
        text_widget.insert(tk.END, f"Erro ao instalar {package}: {e}\n")
    except ValueError as ex:
        text_widget.insert(tk.END, f"Erro ao instalar {package}: Caracteres inválidos no pacote. {ex}\n")


def check_installed_package(requirement):
    try:
        package_name, _, version_spec = requirement.partition("==")
        installed_version = importlib.metadata.version(package_name)
        if version_spec and installed_version != version_spec:
            raise importlib.metadata.PackageNotFoundError
        return True
    except importlib.metadata.PackageNotFoundError:
        return False


def sanitize_requirement(requirement):
    """Remove caracteres nulos e espaços em branco extras."""
    return requirement.replace("\0", "").strip()


def detect_encoding(file_path):
    with open(file_path, 'rb') as file:
        raw_data = file.read()
        result = chardet.detect(raw_data)
        return result['encoding']


def update_requirements(requirements_path, root, text_widget):
    if not Path(requirements_path).is_file():
        text_widget.insert(tk.END, f"Arquivo {requirements_path} não encontrado.\n")
        root.after(2000, root.destroy)
        return

    encoding = detect_encoding(requirements_path)
    with open(requirements_path, 'r', encoding=encoding) as file:
        requirements = file.readlines()

    for requirement in requirements:
        requirement = sanitize_requirement(requirement)
        if not requirement or requirement.startswith("#"):
            continue  # Ignorar linhas vazias ou comentários

        if not check_installed_package(requirement):
            text_widget.insert(tk.END, f"{requirement} não está instalado. Instalando agora...\n")
            root.update()
            install_package(requirement, text_widget)

    text_widget.insert(tk.END, "Instalação concluída.\n")
    root.update()
    root.after(2000, root.destroy)


def main():
    # Inicializar a janela do tkinter
    root = tk.Tk()
    root.title("Instalação de Bibliotecas")
    root.geometry("500x300")

    frame = ttk.Frame(root, padding="10")
    frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(0, weight=1)

    text_widget = tk.Text(frame, wrap="word", height=15)
    text_widget.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    scrollbar = ttk.Scrollbar(frame, orient="vertical", command=text_widget.yview)
    scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
    text_widget['yscrollcommand'] = scrollbar.set

    # Caminho para o arquivo requirements.txt
    requirements_path = Path(__file__).parent.parent.parent.parent / 'requirements.txt'

    # Atualizar requisitos
    root.after(100, update_requirements, requirements_path, root, text_widget)

    root.mainloop()


if __name__ == "__main__":
    main()
