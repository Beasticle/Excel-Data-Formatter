from collections import defaultdict
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
filerList = filterInput.split(', ')
columnList = []
filterListFull = []
for data in range(len(filerList)):
     filterListFull.append(filerList[data].strip(', '))
for data in range(len(colList)):
    columnList.append(colList[data].strip(', '))

filterList = filterInput.split(", ")

for index, data in enumerate(workbook[indexCol]):

    if data == workbook[indexCol].iloc[index]:
        print(columnList)    
        print(data, index)

        # for columnNumber in range(1,len(columnList)):
        #     if workbook[columnList[columnNumber]].iloc

        if workbook[columnList[0]].iloc[index].split('@')[1] not in indexDict and workbook[columnList[0]].iloc[index].split('@')[1] in filterListFull:
            indexDict[workbook[indexCol].iloc[index]] = [workbook[columnList[0]].iloc[index].split('@')[1], workbook[columnList[1]].iloc[index], workbook[columnList[2]].iloc[index]]

        elif workbook[colList[0]].iloc[index].split('@')[1] in filterListFull:
            indexDict[workbook[indexCol].iloc[index]].extend([workbook[columnList[0]].iloc[index].split('@')[1], workbook[columnList[1]].iloc[index], workbook[columnList[2]].iloc[index]])
        else:
            print("Value is already in that list or not valid")

        if index == 7229:
            df = pd.DataFrame.from_dict(indexDict , orient='index')
            df.to_excel("Test.xlsx", sheet_name="Memberoutput")
            print("member list: " + f'{indexDict }')