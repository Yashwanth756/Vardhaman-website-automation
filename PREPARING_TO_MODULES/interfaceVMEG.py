from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import pandas as pd
import threading
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox 
import re
from collections import Counter
from tkinter.ttk import *
import openpyxl
from tkinter import *
from PIL import Image, ImageTk


class gui:
    def __init__(self) -> None:
        root = tk.Tk()
        image = Image.open("images.png")
        photo = ImageTk.PhotoImage(image)
        collegeLabel = tk.Label(image=photo)
        collegeLabel.grid(row=0, column=0, columnspan=2, pady=(0, 50), padx=(0, 0))

        outputLabel = tk.Label(root, text='Output File Name', font=("Arial", 15), bg='white')
        outputLabel.grid(row=1, column=0, padx=(10, 0))

        outputFilename = tk.Entry(borderwidth=3, bg='#E4E9F4', font=('Arial', 15))
        outputFilename.grid(row=1, column=1, columnspan=2, sticky='ew', padx=(0, 10))


        semisterLabel = tk.Label(root, text='Semister', font=("Arial", 15), bg='white')
        semisterLabel.grid(row=2, column=0, padx=(10, 0))

        sem = tk.Entry(borderwidth=3, bg='#E4E9F4', font=('Arial', 15))
        sem.grid(row=2, column=1, columnspan=2, sticky='ew', padx=(0, 10))


        fileBtn = tk.Button(root, text='Choose File', command=openDialog, font=(('Arial', 10)), height=3, bg='gray', fg='black')
        fileBtn.grid(row=3, column=0, pady=(40, 40), sticky='ew', padx=(10, 0))

        self.submit = tk.Button(root, text='Submit', font=('Arial', 10), height=3, bg='green', command=lambda : gatherData(outputFilename.get(), sem.get()))
        self.submit.grid(row=3, column=1, pady=(40, 40), sticky='ew', padx=(10, 10))
        self.submit.configure(state='disabled')


        currentDirectory = tk.Label(root, text='-', font=('Arial', 9), highlightbackground='black', highlightthickness=3, bg='white')

        proper = Frame(root, padx=5, bg='white')

        countPer = tk.Label(proper, text='0%', font=('Arial', 10), bg='white')

        count =  Progressbar(proper, orient = tk.HORIZONTAL, length = 100, mode = 'determinate') 


        count.pack(side=LEFT)
        countPer.pack(side=LEFT)
        root.resizable(False, False)
        root.configure(bg='white')
        root.attributes('-topmost', 1)
        root.mainloop()
    def openDialog(self):
        global filepath
        filepath = filedialog.askopenfilename(initialdir='/', filetypes=(('xl files', '*.xlsx'), ('all files', '*.*')))

        if len(filepath) == 0:
            messagebox.showerror("Error", "No file selected") 
        else:
            self.submit.configure(state='active')
            self.submit.configure(bg='green')
            if len(filepath) > 25:
                self.currentDirectory.configure(text=filepath[:25]+'....')
            else:
                self.currentDirectory.configure(text=filepath)
            selfcurrentDirectory.grid(row=3, column=0, pady=(0, 100),sticky='ew', padx=(10, 0))


    def gatherData(fileName, sem):
        submit.configure(state='disabled')
        fileBtn.configure(state='disabled')
        proper.grid(row=4, column=1, pady=(0, 100), padx=(15, 10), sticky='ew')
        count['value'] = 0
        root.update_idletasks() 
        threading.Thread(target=collectData, args=(filepath, fileName, sem), daemon=True).start()
