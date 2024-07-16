import tkinter as tk

root = tk.Tk()

canvas = tk.Canvas(root, width=1280, height=720, bg="deep sky blue")
canvas.pack()

label = tk.Label(root, text="TASK SCHEDULED BY ELIEZER MORAES SILVA", font=("Arial", 20))
canvas.create_window(1280, 720, window=label)

root.mainloop()
