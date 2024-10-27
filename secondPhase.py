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
