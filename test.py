import tkinter as tk
from tkinter import *
 
class MainWindow:
    def __init__(self, master):
        self.master = master
 
 
        self.frame = tk.Frame(self.master, width = 300, height = 300)
        self.frame.pack()
 
        self.label = tk.Label(self.frame, text = "This is some sample text")
        self.label.place( x = 80, y = 20)
 
        self.button = tk.Button(self.frame, text = "Button")
        self.button.place( x = 120, y = 80)
 
        self.entry = tk.Entry(self.frame)
        self.entry.place( x = 80, y = 160)
        self.btn_var = StringVar()
        self.test()
        self.master.bind("<Key>", lambda event: self.key)
 
    def test(self):
        self.abc = Button(self.master, text="Hello World", command=self.key)
        self.abc.bind("<Key>", lambda event: self.key)
        self.abc.pack(padx=15, pady=20)
    def key(self, event):
        key = event.char
        print(key, 'is pressed')    

    

master = Tk()
root = MainWindow(master)
master.mainloop()