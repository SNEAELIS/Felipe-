import tkinter as tk

window = tk.Tk()
window.geometry("1000x800")
window.title("Hello, Tkinter")
greeting = tk.Label(text="Hello, World!")
greeting.pack()

window.mainloop()