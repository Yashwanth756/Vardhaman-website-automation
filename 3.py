import json
with open('unmodified.json', 'r') as f:
    data = json.load(f)
rollNo = data.keys()
semList=set()
for rn in rollNo:
    semList.update(set(data[rn].keys()))
semList.remove('cgpa')
subList=set()
for rn in rollNo:
    for sem in data[rn].keys():
        if sem == 'cgpa':
            continue
        subList.update(data[rn][sem].keys())
subList.remove('sgpa')

rnc = []
semcol=[]
subcol=[]
gradecol=[]

for rn in rollNo:
    for sem in data[rn].keys():
        if sem =='cgpa':
            continue
        for sub in data[rn][sem].keys():
            if sub=='sgpa':
                continue
            rnc.append(rn)
            semcol.append(sem)
            subcol.append(sub)
            gradecol.append(data[rn][sem][sub])
import pandas as pd
(pd.DataFrame({'rnc':rnc, 'sem':semcol, 'sub':subcol, 'grade':gradecol}).to_excel('completeData.xlsx'))
pd.DataFrame({'sub':subcol}).to_excel('sub.xlsx')
pd.DataFrame({'sem':semcol})