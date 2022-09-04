from sqlite3 import Date
import pandas as pd
from datetime import datetime,date
from docxtpl import DocxTemplate
import os
import glob
import time

print('  ')
print('----**** Welcome ****----')
print(' ')
inp = str(input('Enter the File Path of Folder which contains Candidate Excel Files: '))
print('Note: Remember to put file format(.docx) after the file name')
inp2 = str(input('Enter the File Path of the word template: '))
inp3 = str(input('Enter the File Path to save the Output: '))
    
os.chdir(inp)

for FileList in glob.glob('*.xlsx'):
    df = pd.read_excel('{}'.format(inp)+"\\"+FileList,sheet_name='Salary Fixation Sheet')
    Date = date.today()
    Name = df.iloc[3,2]
    System_title = df.iloc[4,4]
    Business_title = df.iloc[5,4]
    Business_department = df.iloc[6,4]
    Job_level = df.iloc[7,4]
    Job_code = df.iloc[8,4]
    Location = df.iloc[11,4]
    Basic_salary = int(df.iloc[18,3])
    Hra = int(df.iloc[19,3])
    Food_coupouns = int(df.iloc[20,3])
    Travel_assistance = int(df.iloc[21,3])
    Mediclaim = int(df.iloc[22,3])
    Other_allowance = int(df.iloc[23,3])
    Base_salary = int(df.iloc[24,3])
    Pancard = df.iloc[34,2]
    Add = df.iloc[35,2]
    Date_of_joining = df.iloc[36,2]
    # Formatting
    Basic_salary = "{:,}".format(Basic_salary)
    Hra = '{:,}'.format(Hra)
    Food_coupouns = '{:,}'.format(Food_coupouns)
    Travel_assistance = '{:,}'.format(Travel_assistance)
    Mediclaim = '{:,}'.format(Mediclaim)
    Other_allowance = '{:,}'.format(Other_allowance)
    Base_salary1 = Base_salary
    Base_salary = '{:,}'.format(Base_salary)
    # Do not leave the date column blank in the excel file
    Date = Date.strftime('%d-%b-%Y')
    Date_of_joining = Date_of_joining.strftime('%d-%b-%Y')
    #path to save the output
    os.chdir(inp3)
    #zipped = zip(Name,System_title,Business_title,Job_level,Job_code,Location,Basic_salary,hra,Other_allowance,Date_of_joining)
    doc = DocxTemplate(inp2)
    context={'Date':Date,'Name':Name,'System_title':System_title,'Business_title':Business_title,'Business_department':Business_department,'Job_level':Job_level,'Job_code':Job_code,'Location':Location,'Basic_salary':Basic_salary,'Hra':Hra,'Food_coupouns':Food_coupouns,'Travel_assistance':Travel_assistance,'Mediclaim':Mediclaim,'Other_allowance':Other_allowance,'Date_of_joining':Date_of_joining,'Base_salary':Base_salary,'Base_salary1':Base_salary1,'Pancard':Pancard,'Add':Add}
    doc.render(context)
    doc.save('{}.docx'.format(Name))
print(' ')
print('Your files are ready in the Output Folder!!')
print('Output Folder File Path: {}'.format(inp3))
print('--------------------------------------------------------------------------------------------------------------------------')
print('For any queries or suggestions related to software improvement or optimization, contact: robin.kiliyilathu@weareams.com')
time.sleep(10)  