import pandas as pd
import os

if (os.path.exists("Data")):
    pass
else:
    os.mkdir("Data")

settings = dict()
setting_file = open('setting.config', 'r').readlines()
for setting in setting_file:
    [key, value] = setting.rstrip().split('=')
    settings[key] = value

subjects = settings['Subjects'].split(",")
excel = settings['Excel']
excels = []
for row in excel.split("/"):
    column = []
    for col in row.split(","):
        column.append(col)
    excels.append(column)


for i, subject in enumerate(subjects):
    filename = subject.replace(" ", "_").lower() + '.xlsx'
    df = pd.DataFrame([["Roll"] + excels[i]])
    if (os.path.exists('Data/'+filename)):
        pass
    else:
        df.to_excel('Data/'+filename, index=False, header=False)
if (os.path.exists('Data/'+'names.xlsx')):
    pass
else:
    df = pd.DataFrame([["Roll", "Name"]])
    df.to_excel('Data/'+'names.xlsx', index=False, header=False)
