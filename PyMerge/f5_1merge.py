import pandas as pd
import os

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

excelDataFrames = []

tmpDataFrames = []


tableNums = []

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
        matchPos = 'F5_1' in name
   
        names = sourceNames.copy()
        names.remove('F5_1')
        flag = False
        for i in range(0,len(names)):
            matchNeg = names[i] in name
            if(matchNeg):
                flag = True 
        if(matchPos and flag == False):
            filename = name         
    #Название страницы нужной страницы
    if filename != '':
        tmp = pd.read_excel(filename, sheet_name='монографии')
        
        column = tmp['Форма 5_1'].tolist()
        column = list(map(str, column))             
        
        tableNums.clear()
        for j in range(0,len(column)):
            if(column[j].find('I.') != -1 or column[j].find('V.') != -1):
                tableNums.append(j)
                
        tmp = tmp.reset_index(drop=True)
        tmpDataFrames.clear()
        for i in range(0,len(tableNums)):
            tmpDataFrames.append(pd.DataFrame())
            if isFirstFile:
                excelDataFrames.append(pd.DataFrame())
            #Делю датафрейм
            if i==(len(tableNums)-1):
                tmpDataFrames[i] = tmp.loc[tableNums[i]:,:]
                tmpDataFrames[i] = tmpDataFrames[i].reset_index(drop=True)
                tmpDataFrames[i] = tmpDataFrames[i][:-6]
            else:
                tmpDataFrames[i] = tmp.loc[tableNums[i]:(tableNums[i+1]-1),:]
                tmpDataFrames[i] = tmpDataFrames[i].reset_index(drop=True)
                tmpDataFrames[i] = tmpDataFrames[i][:-1]
        for i in range(0,len(tmpDataFrames)):
            if not isFirstFile:
                tmpDataFrames[i] = tmpDataFrames[i].drop([0,1,2,3])
            excelDataFrames[i] = pd.concat([excelDataFrames[i],tmpDataFrames[i]])            
    isFirstFile = False         
    
    

for j in range(0,len(excelDataFrames)):
    ls = []
    ls.extend(range(0, excelDataFrames[j].shape[0]))
    #ls[0] = 0
    for i in range(0,5):
        ls[i] = excelDataFrames[j][excelDataFrames[j].columns[0]][i]
    ls[4] = 1
    k = 2
    for i in range(5,len(ls)):
        ls[i] = k
        k+=1
    excelDataFrames[j][excelDataFrames[j].columns[0]] = ls
    excelData = pd.concat([excelData,excelDataFrames[j]])
        
    


os.chdir(scriptdir)
excelData.to_excel("F5_1_output.xlsx",index=False,header=False,sheet_name='монографии')  

