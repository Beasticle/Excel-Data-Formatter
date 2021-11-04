from collections import defaultdict
from numpy import character
import pandas as pd

print("\n Please put your excel workbook in the same folder as this script. All spelling needs to be exact.")

workbookName = input("\n What is the name of your workbook? ")

sheetName = input("\n What is the name of your sheet? ")

indexCol = input("\n What column is going to be your index? ")

columns = input("\n What columns is your data in? If you have multiple columns, please format them as 'Col1, Col2, etc' ")

filterInput = input("\n What do you want to filter by, eg. scm.uk? ")

ran = False
indexDict = defaultdict(list)
colList = columns.split(", ")
filterList = filterInput.split(', ')
stringToParse = ""
columnList = []
filterListFull = []
filterList = filterInput.split(", ")

workbook = pd.read_excel(f"{workbookName}.xlsx", sheet_name=sheetName)
workbook.head()

print(filterInput)

# for data in range(len(filterList)):
#     print("stripping filters")
#     filterListFull.append(filterList[data].strip(', '))
# for data in range(len(colList)):
#     print("stripping columns")
#     columnList.append(colList[data].strip(', '))

# if filterListFull == []:
#     filterListFull = []
#     for index in range(len(workbook[indexCol])):
#         filterListFull.append(workbook[indexCol].iloc[index])
#         print(filterListFull)
#         if index >= len(workbook[indexCol]):
#             break


def formatFunction():
    for index, data in enumerate(workbook[indexCol]):

        if data == workbook[indexCol].iloc[index] and ran == False:
            print(columnList)    
            print(stringToParse)
            print(filterListFull)
            
            if workbook[stringToParse].iloc[index].split('@')[1] not in indexDict and workbook[stringToParse].iloc[index].split('@')[1] in filterListFull:
                indexDict[workbook[indexCol].iloc[index]] = [workbook[stringToParse].iloc[index].split('@')[1]]
                for i in columnList:
                    print("Formatting, not in dict yet " + workbook[i].iloc[index])
                    indexDict[workbook[indexCol].iloc[index]].extend([workbook[i].iloc[index]])

            elif workbook[stringToParse].iloc[index].split('@')[1] in filterListFull:
                indexDict[workbook[indexCol].iloc[index]].extend([workbook[stringToParse].iloc[index].split('@')[1]])
                for i in columnList:
                    print("Formatting, in dict already " + workbook[i].iloc[index])
                    indexDict[workbook[indexCol].iloc[index]].extend([workbook[i].iloc[index]])
            else:
                print("Value is already in that list or not valid")

            if index == 7229:
                df = pd.DataFrame.from_dict(indexDict , orient='index')
                df.to_excel("Test.xlsx", sheet_name="Memberoutput")
                print("member list: " + f'{indexDict}')
                break

def noFormattingFunction():
    for index, data in enumerate(workbook[indexCol]):

        if data == workbook[indexCol].iloc[index] and ran == False:
            print(columnList)    
            print(data, index)
            print(filterListFull)

            if workbook[string0].iloc[index] not in indexDict and workbook[string0].iloc[index] in filterListFull:
                indexDict[workbook[indexCol].iloc[index]] = []
                for i in columnList:
                    print("No formatting, not in dict yet " + workbook[i].iloc[index])
                    indexDict[workbook[indexCol].iloc[index]].extend([workbook[i].iloc[index]])

            elif workbook[string0].iloc[index] in filterListFull:
                for i in columnList:
                    print("No formatting, in dict already " + workbook[i].iloc[index])
                    indexDict[workbook[indexCol].iloc[index]].extend([workbook[i].iloc[index]])
            else:
                print("Value is already in that list or not valid")

            if index == 7229:
                df = pd.DataFrame.from_dict(indexDict , orient='index')
                df.to_excel("Test.xlsx", sheet_name="Memberoutput")
                print("member list: " + f'{indexDict }')
                break
def noFormattingNoFilter():
    for index, data in enumerate(workbook[indexCol]):

        if data == workbook[indexCol].iloc[index] and ran == False:
            print(columnList)    
            print(data, index)
            print(filterListFull)

            if workbook[string0].iloc[index] not in indexDict: #and workbook[string0].iloc[index] in filterListFull:
                indexDict[workbook[indexCol].iloc[index]] = []
                for i in columnList:
                    print("No formatting, not in dict yet " + workbook[i].iloc[index])
                    indexDict[workbook[indexCol].iloc[index]].extend([workbook[i].iloc[index]])

            # elif workbook[string0].iloc[index]: # in filterListFull:
            #     for i in columnList:
            #         print("No formatting, in dict already " + workbook[i].iloc[index])
            #         indexDict[workbook[indexCol].iloc[index]].extend([workbook[i].iloc[index]])
            else:
                print("Value is already in that list or not valid")

            if index == 7229:
                df = pd.DataFrame.from_dict(indexDict , orient='index')
                df.to_excel("Test.xlsx", sheet_name="Memberoutput")
                print("member list: " + f'{indexDict }')
                break

def formatNoFilter():
    for index, data in enumerate(workbook[indexCol]):

        if data == workbook[indexCol].iloc[index] and ran == False:
            print(columnList)    
            print(stringToParse)
            print(filterListFull)
            
            if workbook[stringToParse].iloc[index].split('@')[1]: # not in indexDict and workbook[stringToParse].iloc[index].split('@')[1] in filterListFull:
                indexDict[workbook[indexCol].iloc[index]] = [workbook[stringToParse].iloc[index].split('@')[1]]
                for i in columnList:
                    print("Formatting, not in dict yet " + workbook[i].iloc[index])
                    indexDict[workbook[indexCol].iloc[index]].extend([workbook[i].iloc[index]])

            # elif workbook[stringToParse].iloc[index].split('@')[1] in filterListFull:
            #     indexDict[workbook[indexCol].iloc[index]].extend([workbook[stringToParse].iloc[index].split('@')[1]])
            #     for i in columnList:
            #         print("Formatting, in dict already " + workbook[i].iloc[index])
            #         indexDict[workbook[indexCol].iloc[index]].extend([workbook[i].iloc[index]])
            else:
                print("Value is already in that list or not valid")

            if index == 7229:
                df = pd.DataFrame.from_dict(indexDict , orient='index')
                df.to_excel("Test.xlsx", sheet_name="Memberoutput")
                print("member list: " + f'{indexDict}')
                break


for col in columnList:
    if '@' not in [workbook[col].iloc[0]]:
        if filterListFull == []:
            for i in range(len(workbook[col])):
                #print(i)
                globals()[f"string{i}"] = col
                if i >= len(columnList):
                    noFormattingNoFilter()
                    ran = True
                    break
        else:
            for i in range(len(workbook[col])):
                #print(i)
                globals()[f"string{i}"] = col
                if i >= len(columnList):
                    noFormattingFunction()
                    ran = True
                    break
    else: 
        if '@' in [workbook[col].iloc[0]]:
            stringToParse = col
            ran = True
        else:
            if filterListFull == []:
                for i in enumerate(workbook[col], start=1):
                    if i == 1:
                        string0 = col
                    elif i >= len(columnList):
                        formatNoFilter()
                        ran = True
                        break
                    else:
                        globals()[f"string{i}"] = col
            else:
                for i in enumerate(workbook[col], start=1):
                    if i == 1:
                        string0 = col
                    elif i >= len(columnList):
                        formatFunction()
                        ran = True
                        break
                    else:
                        globals()[f"string{i}"] = col