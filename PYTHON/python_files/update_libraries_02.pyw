import subprocess
import sys

import pkg_resources
import tkinter as tk
from tkinter import messagebox


def check_and_update_requirements():
    try:
        # Ler o arquivo requirements.txt
        with open('requirements.txt', 'r') as file:
            requirements = file.readlines()

        # Verificar e atualizar bibliotecas
        installed_packages = {pkg.key: pkg.version for pkg in pkg_resources.working_set}
        packages_to_update = []

        for requirement in requirements:
            requirement = requirement.strip()
            if requirement and requirement.split('==')[0] not in installed_packages:
                packages_to_update.append(requirement)

        if packages_to_update:
            # Exibir janela de status
            root = tk.Tk()
            root.withdraw()  # Esconde a janela principal
            messagebox.showinfo("Atualizando Bibliotecas",
                                f"Atualizando as seguintes bibliotecas:\n{', '.join(packages_to_update)}")

            # Atualizar bibliotecas
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--upgrade', *packages_to_update])

            messagebox.showinfo("Atualização Concluída", "As bibliotecas foram atualizadas com sucesso.")
            root.destroy()

        else:
            print("Todas as bibliotecas já estão atualizadas.")
    except Exception as e:
        print(f"Erro ao atualizar bibliotecas: {e}")


def main():
    check_and_update_requirements()

    # Executar o script principal
    # Coloque aqui o código do seu script principal
    print("Executando o script principal...")


if __name__ == "__main__":
    main()
