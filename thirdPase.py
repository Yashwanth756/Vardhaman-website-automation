import json
import pandas as pd
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
# OutputFormating(3)