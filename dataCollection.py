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


class CollectData:
    def __init__(self) -> None:    
        self.backlogsCode = []
        self.backlogsName = []
    
    def marks(self, rollNo, password, sem):
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