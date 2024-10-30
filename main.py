from firstPhase import *
from secondPhase import *
from thirdPase import *
from tkinter.ttk import *
from tkinter import *
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox 
def openDialog():
    global filepath
    filepath = filedialog.askopenfilename(initialdir='/', filetypes=(('xl files', '*.xlsx'), ('all files', '*.*')))

    if len(filepath) == 0:
        messagebox.showerror("Error", "No file selected") 
    else:
        submit.configure(state='active')
        submit.configure(bg='green')
        if len(filepath) > 25:
            currentDirectory.configure(text=filepath[:25]+'....')
        else:
            currentDirectory.configure(text=filepath)
        currentDirectory.grid(row=4, column=0, pady=(0, 100),sticky='ew', padx=(10, 0))

def gatherData(fileName, sem, n_chunks):
    submit.configure(state='disabled')
    fileBtn.configure(state='disabled')
    root.update_idletasks() 
    threading.Thread(target=process, args=(filepath, fileName, int(sem), int(n_chunks)), daemon=True).start()

def process(filepath, fileName, sem, n_chunks=3):
    data = DataCollection(n_chunks, filepath).divdata
    DataProcessing(data, n_chunks)
    OutputFormating(n_chunks, fileName, sem)
    print(filepath, fileName, sem)

if __name__=="__main__":
    # n_chunks = 3
    # data = DataCollection(n_chunks).divdata
    # DataProcessing(data, n_chunks)
    # OutputFormating(3)



    root = tk.Tk()
    image = Image.open("images.png")
    photo = ImageTk.PhotoImage(image)
    collegeLabel = tk.Label(image=photo)
    collegeLabel.grid(row=0, column=0, columnspan=2, pady=(0, 50), padx=(0, 0))

    outputLabel = tk.Label(root, text='Output File Name', font=("Arial", 15), bg='white')
    outputLabel.grid(row=1, column=0, padx=(10, 0))

    outputFilename = tk.Entry(borderwidth=3, bg='#E4E9F4', font=('Arial', 15))
    outputFilename.grid(row=1, column=1, columnspan=2, sticky='ew', padx=(0, 10))


    semisterLabel = tk.Label(root, text='SEMESTER', font=("Arial", 15), bg='white')
    semisterLabel.grid(row=2, column=0, padx=(10, 0))

    sem = tk.Entry(borderwidth=3, bg='#E4E9F4', font=('Arial', 15))
    sem.grid(row=2, column=1, columnspan=2, sticky='ew', padx=(0, 10))

    tabLabel = tk.Label(root, text='TABS', font=("Arial", 15), bg='white')
    tabLabel.grid(row=3, column=0, padx=(10, 0))

    tab = tk.Entry(borderwidth=3, bg='#E4E9F4', font=('Arial', 15))
    tab.grid(row=3, column=1, columnspan=2, sticky='ew', padx=(0, 10))





    fileBtn = tk.Button(root, text='Choose File', command=openDialog, font=(('Arial', 10)), height=3, bg='gray', fg='black')
    fileBtn.grid(row=4, column=0, pady=(40, 40), sticky='ew', padx=(10, 0))

    submit = tk.Button(root, text='Submit', font=('Arial', 10), height=3, bg='green', command=lambda : gatherData(outputFilename.get(), sem.get(), tab.get()))
    submit.grid(row=4, column=1, pady=(40, 40), sticky='ew', padx=(10, 10))
    submit.configure(state='disabled')


    currentDirectory = tk.Label(root, text='-', font=('Arial', 9), highlightbackground='black', highlightthickness=3, bg='white')

    root.resizable(False, False)
    root.configure(bg='white')
    root.attributes('-topmost', 1)
    root.mainloop()
