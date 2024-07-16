# script.py
import time

def task1():
    time.sleep(1)  # Simula uma tarefa demorada
    return 10

def task2():
    time.sleep(1)
    return 20

def task3():
    time.sleep(1)
    return 30

def run_tasks(progress_callback):
    total_steps = 3
    progress_callback(1 / total_steps * 100)
    task1()
    
    progress_callback(2 / total_steps * 100)
    task2()
    
    progress_callback(3 / total_steps * 100)
    task3()
    
    progress_callback(100)

if __name__ == "__main__":
    run_tasks(print)  # Use print para debug se rodar este script diretamente
