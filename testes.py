import tkinter as tk

root = tk.Tk()


def toggle():
    ninja_button.pack()


ninja_button = tk.Button(root, text='Aha!')
visible_button = tk.Button(root, text='Show', command=toggle)
visible_button.pack()

root.mainloop()