from collections import defaultdict
from numpy import radians
import pandas as pd

print("\n Please put your excel workbook in the same folder as this script. All spelling needs to be exact.")

workbookName = input("\n What is the name of your workbook? ")

sheetName = input("\n What is the name of your sheet? ")

indexCol = input("\n What column is going to be your index? ")

columns = input("\n What columns is your data in? If you have multiple columns, please format them as 'Col1, Col2, etc' ")

filterInput = input("\n What do you want to filter by, eg. scm.uk? ")

workbook = pd.read_excel(f"{workbookName}.xlsx", sheet_name=sheetName)
workbook.head()

ran = False
indexDict = defaultdict(list)
colList = columns.split(", ")
filterList = filterInput.split(', ')
stringToParse = ""
columnList = []
filterListFull = []
for data in range(len(filterList)):
     filterListFull.append(filterList[data].strip(', '))
for data in range(len(colList)):
    columnList.append(colList[data].strip(', '))

filterList = filterInput.split(", ")

if filterListFull == []:
    for index in range(len(workbook[indexCol])):
        filterListFull.append(workbook[indexCol].iloc[index])


def formatFunction():
    for index, data in enumerate(workbook[indexCol]):

        if data == workbook[indexCol].iloc[index] and ran == False:
            print(columnList)    
            print(data, index)
            print(filterListFull)
            
            if workbook[stringToParse].iloc[index].split('@')[1] not in indexDict and workbook[stringToParse].iloc[index].split('@')[1] in filterListFull:
                indexDict[workbook[indexCol].iloc[index]] = [workbook[stringToParse].iloc[index].split('@')[1]]
                for i in columnList:
                    #if i not in stringToParse:
                    indexDict[workbook[indexCol].iloc[index]].extend([workbook[i].iloc[index]])

            elif workbook[stringToParse].iloc[index].split('@')[1] in filterListFull:
                indexDict[workbook[indexCol].iloc[index]].extend([workbook[stringToParse].iloc[index].split('@')[1]])
                for i in columnList:
                    #if i not in stringToParse:
                    indexDict[workbook[indexCol].iloc[index]].extend([workbook[i].iloc[index]])
            else:
                print("Value is already in that list or not valid")

            if index == 7229:
                df = pd.DataFrame.from_dict(indexDict , orient='index')
                df.to_excel("Test.xlsx", sheet_name="Memberoutput")
                print("member list: " + f'{indexDict }')
                break
                exit

def noFormattingFunction():
    for index, data in enumerate(workbook[indexCol]):

        if data == workbook[indexCol].iloc[index] and ran == False:
            print(columnList)    
            print(data, index)
            print(filterListFull)

            if workbook[string0].iloc[index] not in indexDict and workbook[string0].iloc[index] in filterListFull:
                indexDict[workbook[indexCol].iloc[index]] = []
                for i in columnList:
                    indexDict[workbook[indexCol].iloc[index]].extend([workbook[i].iloc[index]])

            elif workbook[string0].iloc[index] in filterListFull:
                for i in columnList:
                    indexDict[workbook[indexCol].iloc[index]].extend([workbook[i].iloc[index]])
            else:
                print("Value is already in that list or not valid")

            if index == 7229:
                df = pd.DataFrame.from_dict(indexDict , orient='index')
                df.to_excel("Test.xlsx", sheet_name="Memberoutput")
                print("member list: " + f'{indexDict }')
                break
                exit


for col in columnList:
    if str('@') not in [workbook[col].iloc[0]]:
        for i in range(len(workbook[col])):
            print(i)
            globals()[f"string{i}"] = col
            noFormat = True
            ran = True
    else: 
        if str('@') in [workbook[col].iloc[0]]:
            noFormat = False
            stringToParse = col
            ran = True
        else:
            for i in enumerate(workbook[col], start=1):
                globals()[f"string{i}"] = col
                ran = True

if noFormat:
    noFormattingFunction()
else:
    formatFunction()