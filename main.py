from collections import defaultdict
import pandas as pd

print("Please put your excel workbook in the same folder as this script. All spelling needs to be exact.")

sheetName = input("What is the name of your sheet?")

indexCol = input('What column is going to be your index?')

columns = input("What columns is your data in? If you have multiple columns, please format them as 'Col1, Col2, etc'")

workbook = pd.read_excel("*.xlsx", sheet_name=sheetName)
workbook.head()

indexDict = defaultdict(list)
dataColList = columns.split(", ")

for index, data in enumerate(workbook['Member Mail']):

    if data == workbook['Member Mail'].iloc[index]:

        print(data + " " + workbook['Member Mail'].iloc[index].split('@')[1] + f"\n")

        if workbook['Member Mail'].iloc[index].split('@')[1] not in indexDict and workbook['Member Mail'].iloc[index].split('@')[1] in dataColList: # or 'smc.uk':
            indexDict[workbook[indexCol].iloc[index]] = [workbook['Member Mail'].iloc[index].split('@')[1], workbook['Teams Name'].iloc[index], workbook['Role'].iloc[index]]

        elif workbook['Member Mail'].iloc[index].split('@')[1] in dataColList: # or 'smc.uk':
            indexDict[workbook[indexCol].iloc[index]].extend([workbook['Member Mail'].iloc[index].split('@')[1], workbook['Teams Name'].iloc[index], workbook['Role'].iloc[index]])
        else:
            print("Value is already in that list or not valid")

        if index == 7229:
            df = pd.DataFrame.from_dict(indexDict , orient='index')
            df.to_excel("/Output/Test.xlsx", sheet_name="Memberoutput")
            print("member list: " + f'{indexDict }')