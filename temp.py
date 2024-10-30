import json
with open('22881A7356.json', 'r') as f:
    data = json.load(f)
print(data['Semester - I']['22881A7302'].keys())
# with open('sem4.json', 'w') as f:
#     json.dump(data['Semester - I'], f, indent=4)