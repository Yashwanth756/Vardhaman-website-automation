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
        with open(outputfile+'.json', 'w') as f:
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
        finalData.to_excel(outputfile+'.xlsx')

if __name__=="__main__":
    n_chunks = 3
    data = DataCollection(n_chunks).divdata
    DataProcessing(data, n_chunks)
    OutputFormating(3)


