import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import threading
import re
import openpyxl
import os
import json
import numpy as np
from firstPhase import DataCollection
import multiprocessing
import json
from tkinter.ttk import *
from tkinter import *
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox 

class DataCollection:
    def __init__(self, n_chunks, filePath='data.xlsx') -> None:
        file_path = filePath
        data = pd.read_excel(file_path)
        scannedData = self.scanData(data)
        self.divdata = self.divideData(data)

        # Making chunks and resetting index
        for ind, data in enumerate(self.divdata):
            data = self.chunks(data, n_chunks)
            # Reset index for each chunk
            self.divdata[ind] = [chunk.reset_index(drop=True) for chunk in data if not chunk.empty]

    def chunks(self, data, n):
        return np.array_split(data, n)

    def scanData(self, data):
        return data 
    def divideData(self, data):
        return [data]  
    

class DataProcessing:
    def __init__(self,data, n_chunks=0) -> None:
        # data = DataCollection(n_chunks).divdata
        processes = []
        id=0
        for chunk in data[0]:
            print((chunk), '\n\n')
            process = multiprocessing.Process(target=MarksRecord, args=(id, chunk.copy(), ))
            id+=1
            processes.append(process)
            process.start()
        for process in processes:
            process.join()
        print('all completed')

class MarksRecord:
    def __init__(self, id, data=None) -> None:
        print('Initiating driver....')
        driver = webdriver.Chrome()
        self.classRecords={}
        self.failedRecords={}
        for i in range(len(data)):
                if len(data['rollNo'][i]) == 10:
                    try:
                        rollNo = data['rollNo'][i]
                        password = data['password'][i]
                        scoreCard=self.studentRecord(driver, data['rollNo'][i], data['password'][i])
                        self.classRecords[rollNo]=scoreCard
                    except:
                        self.failedRecords[rollNo]=password
        self.saveData(self.classRecords, id)
        self.saveData(self.failedRecords,'failed_' +str(id))
        driver.quit()  # Close the driver after processing
   
    def studentRecord(self, driver, rollNo, password):
        score = {}
        semList = ['Semester - I', 'Semester - II', 'Semester - III', 'Semester - IV', 'Semester - V', 'Semester - VI', 'Semester - VII', 'Semester - VIII']
        semCount=0
        driver.get('https://studentscorner.vardhaman.org/')
        driver.implicitly_wait(5)
        driver.find_element(By.ID, 'username').send_keys(rollNo)
        password = driver.find_element(By.ID, 'login-pass').send_keys(password)
        print('loged in : ', rollNo)
        driver.find_element(By.NAME, 'ok').click()

        driver.implicitly_wait(2)
        marks_link = driver.find_element(By.LINK_TEXT, 'Credit Register').click()
        subGrade={}
        semGrade={}
        data = driver.find_elements(By.TAG_NAME, 'tr')
        semData = data[10:-5]
        overallData = data[-4:-2:1]
        self.c=0
        for tr in semData:
            info = tr.get_attribute('innerHTML')
            tags = re.findall('>.+<', info)

            if len(tags)==7:#identifies only subject
                name=tags[2][1:-1]
                grade=tags[3][1:-1]
                subGrade[name]=grade
            else:
                if 'Semester Grade Point Average' in tags[0]:
                            subGrade['sgpa']= tags[0].split(': ')[-1][:3]
                else:
                    for sem in semList:#adding to sem
                        if sem == tags[0][2:-1]:
                            semGrade[semList[self.c]]=subGrade.copy()
                            subGrade={}
                            self.c+=1
                            break
                    
        semGrade[semList[self.c]]=subGrade.copy()
        semGrade['cgpa']=overallData[-1].get_attribute('innerHTML').split(': ')[-1][:3]
        return semGrade

    def saveData(self, data, filename='data.json'):
        file_path = str(filename) + '.json'
        # os.makedirs(os.path.dirname(file_path), exist_ok=True)
        with open(file_path, "w") as json_file:
            json.dump(data, json_file, indent=4)

class OutputFormating:
    def __init__(self, n_chunks, outputfile='formattedData', semester=None) -> None:
        data = {}
        for i in range(0, n_chunks):
            with open(str(i)+'.json', 'r') as f:
                r=json.load(f)
                data.update(r)
        rollNo = data.keys()
        semList = data[list(rollNo)[0]].keys()
        self.fmtdata={}
        for sem in semList:
            self.fmtdata[sem]={}
            for rn in rollNo:
                sd = {}
                subGrade = data[rn][sem]
                sd[rn]=subGrade
                self.fmtdata[sem].update(sd)
        pd.DataFrame(self.fmtdata).to_excel(outputfile+'.xlsx')
        with open('22881A7356'+'.json', 'w') as f:
            json.dump(self.fmtdata, f, indent=4)
        self.semData(outputfile, semester)
        # print(nd)
    def semData(self,outputfile, semester=None):
        print(type(semester))
        semList=list(self.fmtdata.keys())
        # print(semList)
        if semester == None:
            data=self.fmtdata[semList[-2]]
        elif (len(semList)-1)>=semester :
            data=self.fmtdata[semList[semester-1]]
        else:
            return
        
        rollNo = list(data.keys())
        subjects = list(data[rollNo[0]].keys())

        allSub=[]
        print(subjects)
        for sub in subjects:
            score=[]
            for rn in rollNo:
                score.append(data[rn][sub])
            allSub.append(score)
        print(allSub)
        finalData=pd.DataFrame()
        finalData['rollNo']=rollNo
        for ind, sub in enumerate(allSub):
            finalData[subjects[ind]]=sub
        print(finalData)
        finalData.to_excel('Results.xlsx')
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


