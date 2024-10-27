import json
import pandas as pd
data = {}
for i in range(0, 3):
    with open(str(i)+'.json', 'r') as f:
        r=json.load(f)
        data.update(r)
rollNo = data.keys()
semList = data[list(rollNo)[0]].keys()
nd={}
for sem in semList:
    nd[sem]={}
    for rn in rollNo:
        sd = {}
        subGrade = data[rn][sem]
        sd[rn]=subGrade
        nd[sem].update(sd)
(pd.DataFrame(nd).to_excel('final.xlsx'))
with open('output.json', 'w') as f:
    json.dump(nd, f, indent=4)
print(nd)