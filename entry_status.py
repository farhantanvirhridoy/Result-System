import pandas as pd


settings = dict()
setting_file = open('setting.config', 'r').readlines()
for setting in setting_file:
    [key, value] = setting.rstrip().split('=')
    settings[key] = value

response=[]

subject_file_names = settings['Subjects'].split(",")
for subject_file_name in subject_file_names:
    filename = subject_file_name.replace(" ", "_").lower()
    exec(f"{filename} = pd.read_excel('data/{filename}.xlsx')")
    exec(f"count_nan = {filename}.isna().sum().sum()")
    if (count_nan != 0):

        response.append(str(count_nan) + " data missing in " + subject_file_name)
    else:
        response.append("All data is present in " + subject_file_name)

names = pd.read_excel('Data/names.xlsx')
count_nan = names.isna().sum().sum()
if (count_nan != 0):
    response.append(str(count_nan) + " data missing in " + 'Names')
else:
    response.append("All data is present in " + 'Names')

def status():
    return response


