import pandas as pd
import os
import re

scriptdir = os.getcwd()
os.chdir(scriptdir)
folderNames = os.listdir(path=".")

del_list = []

for name in folderNames:
    if name.endswith('.py') or name.endswith('.xlsx') or name.endswith('.ipynb_checkpoints'):      
        del_list.append(name)
for name in del_list:
    folderNames.remove(name)


isFirstFile = True

#Все имена форм
sourceNames = ['F3','F3_1','F4_1','F4','F5_1','F5_2(0)','F5_2(1)',
         'F5_2(2)','F5_3','F5_4','F5_5','F5_6','F5_7',
         'F6_1','F6_2','F6','F7','F7_1','F8','F9_1','F10','F11','F12']



names =  sourceNames.copy()


excelData = pd.DataFrame()

flag = False

#Перебираем директории, которые располагаются рядом со скриптом
for dir in folderNames:
    #print(dir)
    os.chdir(scriptdir)
    os.chdir(scriptdir + "/" + dir)
    fileNames = os.listdir(path=".")
    filename = '';
    names =  sourceNames.copy()
    for name in fileNames:
        #Название форм, с которыми сейчас работаем
        matchPos = re.search('F3_1', name)
        names = sourceNames.copy()
        names.remove('F3')
        names.remove('F3_1')
        flag = False
        for i in range(0,len(names)):
            matchNeg = re.search(names[i], name)
            if(matchNeg != None):
                flag = True 
        if(matchPos != None and flag == False):
            filename = name
    #Название нужной страницы
    if filename != '':
        tmp = pd.read_excel(filename, sheet_name='поданные проекты')
    # Удаляю первые строки
        if isFirstFile:
            tmp = tmp.drop([0,1,2,3])
            isFirstFile = False
        else:
            #Удаляю первые ВМЕСТЕ С ШАПКОЙ
            tmp = tmp.drop([0,1,2,3,4,5]) 
        #Удаляю последние 
        tmp = tmp.loc[0:(tmp[tmp.isnull().all(axis=1)].index[0]-1),:]
        excelData = pd.concat([excelData,tmp])

ls = []
ls.extend(range(0, excelData.shape[0]))

ls[0] = 0
for i in range(2,3):
    ls[i] = 1

j = 2
for i in range(3,len(ls)):
    ls[i] = j
    j+=1

excelData = excelData.reset_index(drop=True)
pp = excelData[excelData.columns[0]][0]
excelData[excelData.columns[0]] = ls
excelData[excelData.columns[0]][0] = pp

os.chdir(scriptdir)
excelData.to_excel("F3_1_output.xlsx",index=False,header=False,sheet_name='поданные проекты')  


