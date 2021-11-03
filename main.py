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

indexDict = defaultdict(list)
colList = columns.split(", ")
filterList = filterInput.split(', ')
stringToParse = ""
# string1 = ""
# string2 = ""
columnList = []
filterListFull = []
for data in range(len(filterList)):
     filterListFull.append(filterList[data].strip(', '))
for data in range(len(colList)):
    columnList.append(colList[data].strip(', '))

for data in columnList:
    if "@" not in workbook[data].iloc[0]:
        for i in range(len(workbook[data])):
            globals()[f"string{i}"] = data
            noformat = True
    else: 
        if "@" in workbook[data].iloc[0]:
            stringToParse = data
            noformat = False
        else:
            for i in enumerate(workbook[data], start=1):
                globals()[f"string{i}"] = data

filterList = filterInput.split(", ")

def formatFunction():

    for index, data in enumerate(workbook[indexCol]):

        if data == workbook[indexCol].iloc[index]:
            print(columnList)    
            print(data, index)

            if workbook[stringToParse].iloc[index].split('@')[1] not in indexDict and workbook[stringToParse].iloc[index].split('@')[1] in filterListFull:
                indexDict[workbook[indexCol].iloc[index]] = [workbook[stringToParse].iloc[index].split('@')[1], workbook[string1].iloc[index], workbook[string2].iloc[index]]

            elif workbook[stringToParse].iloc[index].split('@')[1] in filterListFull:
                indexDict[workbook[indexCol].iloc[index]].extend([workbook[stringToParse].iloc[index].split('@')[1], workbook[string1].iloc[index], workbook[string2].iloc[index]])
            else:
                print("Value is already in that list or not valid")

            if index == 7229:
                df = pd.DataFrame.from_dict(indexDict , orient='index')
                df.to_excel("Test.xlsx", sheet_name="Memberoutput")
                print("member list: " + f'{indexDict }')

def noFormattingFunction():
    for index, data in enumerate(workbook[indexCol]):

        if data == workbook[indexCol].iloc[index]:
            print(columnList)    
            print(data, index)

            if workbook[string0].iloc[index] not in indexDict and workbook[string0].iloc[index] in filterListFull:
                indexDict[workbook[indexCol].iloc[index]] = [workbook[string0].iloc[index], workbook[string1].iloc[index], workbook[string2].iloc[index]]

            elif workbook[string0].iloc[index] in filterListFull:
                indexDict[workbook[indexCol].iloc[index]].extend([workbook[string0].iloc[index], workbook[string1].iloc[index], workbook[string2].iloc[index]])
            else:
                print("Value is already in that list or not valid")

            if index == 7229:
                df = pd.DataFrame.from_dict(indexDict , orient='index')
                df.to_excel("Test.xlsx", sheet_name="Memberoutput")
                print("member list: " + f'{indexDict }')

if noformat == True:
    formatFunction()
else:
    noFormattingFunction()