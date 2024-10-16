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

backlogsCode = []
backlogsName = []

def marks(rollNo, password):
    score = []
    score.append(rollNo)
    driver.get('https://studentscorner.vardhaman.org/')
    driver.implicitly_wait(5)
    driver.find_element(By.ID, 'username').send_keys(rollNo)
    password = driver.find_element(By.ID, 'login-pass').send_keys(password)
    print('loged in : ', rollNo)
    driver.find_element(By.NAME, 'ok').click()

    driver.implicitly_wait(2)
    marks_link = driver.find_element(By.LINK_TEXT, 'Credit Register').click()

    for n in driver.find_elements(By.TAG_NAME, 'th'):
        if 'Semester Grade Point Average' in str(n.text)   or 'Cumulative Grade Point Average' in  str(n.text) :
            score.append(float((n.text).split(':')[-1]))
    backlogs = driver.find_elements(By.CSS_SELECTOR, "[bgcolor ='pink']")
    subjectCode = []
    subjectName = []
    if len(backlogs) > 0:
        for i in range(len(backlogs)):
            s = backlogs[i].get_attribute('innerHTML')
            tags = re.findall('<.+>', s)
            code = re.findall('>.+<', tags[1])[0][1:-1]
            subjectCode.append(code)
            backlogsCode.append(code)
            name = re.findall('>.+<', tags[2])[0][1:-1]
            backlogsName.append(name)
            subjectName.append(name)
            print(tags)
    score.append(len(backlogs))
    score.append(subjectCode)
    score.append(subjectName)
    print(score, '\n\n')
    return score

def collectData(filepath, fileName):
    global driver
    print('Initiating driver....')
    driver = webdriver.Chrome()
    studentScore = []
    print('Reading data\n\n')
    data = pd.read_excel(filepath)
    n = data.shape[0]
    for i in range(len(data)):
        studentScore.append(marks(data['rollNo'][i], data['password'][i]))
        value = int(((i+1)/n)*100)
        count['value'] = value
        root.update_idletasks() 
    colLen = len(studentScore[0])
    columns = ['sgpa '+str(i+1) for i in range(colLen-5)]
    columns.insert(0, 'rollNo')
    columns.append('cgpa')
    columns.append('backlogs')
    columns.append('subjectsCode')
    columns.append('subjectsName')

    data = pd.DataFrame(columns=columns, data=studentScore)

    if (len(fileName) == 0):
        name = 'StudentMarks'
    else:
        name = fileName
    data.to_excel(name +'.xlsx', index=False)

    driver.quit()
    fileBtn.configure(state='active')
    bc = dict(Counter(backlogsCode))
    bn = dict(Counter(backlogsName))
    print('Summary of backlogs(code)\n')
    print(bc, '\n\n')
    print('Summary of backlogs(name)\n')
    print(bn, '\n\n')
    for key, value in bc.items():
        bc[key]= [value]
    for key, value in bn.items():
        bn[key]= [value]
    pd.DataFrame(index= bc.keys(), data= bc.values()).to_excel(name + '_backlogs_code_list.xlsx')
    pd.DataFrame(index= bn.keys(), data= bn.values()).to_excel(name + '_backlogs_subjectName_list.xlsx')


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
        currentDirectory.grid(row=3, column=0, columnspan=2, pady=(0, 100),sticky='ew', padx=(10, 0))


def gatherData(fileName):
    submit.configure(state='disabled')
    fileBtn.configure(state='disabled')
    count.grid(row=3, column=2, pady=(0, 100), padx=(10, 10), sticky='ew', ipadx=5)
    count['value'] = 0
    root.update_idletasks() 
    threading.Thread(target=collectData, args=(filepath, fileName), daemon=True).start()


root = tk.Tk()
collegeLabel = tk.Label(root, text='Vardhaman College of engineering', font=(('Arial', 20)), bg='gray')
collegeLabel.grid(row=0, column=0, columnspan=3, pady=(100, 50), padx=(10, 10))

outputLabel = tk.Label(root, text='output file name', font=("Arial", 15), bg='gray')
outputLabel.grid(row=1, column=0, padx=(10, 0))

outputFilename = tk.Entry(borderwidth=3, bg='#E4E9F4', font=('Arial', 15))
outputFilename.grid(row=1, column=1, columnspan=2, sticky='ew', padx=(0, 10))


fileBtn = tk.Button(root, text='choose file', command=openDialog, font=(('Arial', 10)), height=3, bg='gray', fg='black')
fileBtn.grid(row=2, column=0, pady=(40, 40), sticky='ew', padx=(10, 0))

submit = tk.Button(root, text='Submit', font=('Arial', 10), height=3, bg='green', command=lambda : gatherData(outputFilename.get()))
submit.grid(row=2, column=1, columnspan=2, pady=(40, 40), sticky='ew', padx=(10, 10))
submit.configure(state='disabled')


currentDirectory = tk.Label(root, text='-', font=('Arial', 15), highlightbackground='gray', highlightthickness=2, bg='gray')


count =  Progressbar(root, orient = tk.HORIZONTAL, length = 100, mode = 'determinate') 



root.resizable(False, False)
root.configure(bg='gray')
root.attributes('-topmost', 1)
root.mainloop()
