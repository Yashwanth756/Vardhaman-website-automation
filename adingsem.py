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
backlogsCode = []
backlogsName = []

def marks(rollNo, password, sem):
    print(sem)
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
    flag=0
    subjectGrade=[]
    for tr in driver.find_elements(By.TAG_NAME, 'tr')[1:]:
        td = tr.get_attribute('innerHTML')
        tags = re.findall('>.+<', td)
        if flag:
            if 'Semester' in td:
                break
            if len(tags)>3:
                subjectGrade.append(tags[2][1:-1]+' : ' + tags[3][1:-1])

        for i in tags:
            if 'Semester - '+sem in i:
                flag=1
    # print(subjectGrade)

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
    score.append(len(backlogs))
    score.append(subjectCode)
    score.append(subjectName)
    score.append(subjectGrade)
    print(score, '\n\n')
    return score

def collectData(filepath, fileName, sem):
    global driver
    print('Initiating driver....')
    driver = webdriver.Chrome()
    studentScore = []
    print('Reading data\n\n')
    data = pd.read_excel(filepath)
    n = data.shape[0]
    failedData=[]
    for i in range(len(data)):
        if len(data['rollNo'][i]) == 10:
            try:
                studentScore.append(marks(data['rollNo'][i], data['password'][i], sem))
            except:
                value = int(((i+1)/n)*100)
                countPer.configure(text=str(value)+'%')
                count['value'] = value
                root.update_idletasks() 
                failedData.append(data['rollNo'][i])
                continue
            value = int(((i+1)/n)*100)
            countPer.configure(text=str(value)+'%')
            count['value'] = value
            root.update_idletasks() 

    colLen = len(studentScore[0])
    columns = ['sgpa '+str(i+1) for i in range(colLen-6)]
    columns.insert(0, 'rollNo')
    columns.append('cgpa')
    columns.append('backlogs')
    columns.append('subjectsCode')
    columns.append('subjectsName')
    columns.append('subjectGrade')

    data = pd.DataFrame(columns=columns, data=studentScore)
    if len(failedData)>0:
        pd.DataFrame(failedData).to_excel('failed.xlsx')
    else:
        pd.DataFrame(['none']).to_excel('failed.xlsx')
    if (len(fileName) == 0):
        name = 'StudentMarks'
    else:
        name = fileName
    
    data.to_excel(name + 'Sem-'+sem+ '.xlsx', index=False)

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

    pd.DataFrame(index= bc.keys(), data= bc.values()).to_excel(name +'Sem-'+sem+ '_backlogs_code_list.xlsx')
    pd.DataFrame(index= bn.keys(), data= bn.values()).to_excel(name +'Sem-'+sem+'_backlogs_subjectName_list.xlsx')


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
        currentDirectory.grid(row=3, column=0, pady=(0, 100),sticky='ew', padx=(10, 0))


def gatherData(fileName, sem):
    submit.configure(state='disabled')
    fileBtn.configure(state='disabled')
    proper.grid(row=4, column=1, pady=(0, 100), padx=(15, 10), sticky='ew')
    count['value'] = 0
    root.update_idletasks() 
    threading.Thread(target=collectData, args=(filepath, fileName, sem), daemon=True).start()


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

submit = tk.Button(root, text='Submit', font=('Arial', 10), height=3, bg='green', command=lambda : gatherData(outputFilename.get(), sem.get()))
submit.grid(row=3, column=1, pady=(40, 40), sticky='ew', padx=(10, 10))
submit.configure(state='disabled')


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
