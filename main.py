from collections import defaultdict
import pandas as pd

workbook = pd.read_excel("EMEATeamsPermissions.xlsx", sheet_name="EMEATeamsPermissions")
workbook.head()

memberDict = defaultdict(list)

for index, domainName in enumerate(workbook['Member Mail']):
    print(domainName + f" {index}")

    if domainName == workbook['Member Mail'].iloc[index]:

        print(domainName + " " + workbook['Member Mail'].iloc[index].split('@')[1] + f"\n")

        if workbook['Member Mail'].iloc[index].split('@')[1] not in memberDict:
            memberDict[workbook['Member Mail'].iloc[index].split('@')[1]] = [workbook['Member Name'].iloc[index]]
            if workbook['Member Name'].iloc[index] not in memberDict[workbook['Member Mail'].iloc[index].split('@')[1]]:
                memberDict[workbook['Member Mail'].iloc[index].split('@')[1]].append(workbook['Member Name'].iloc[index])
            else:
                print("Value is already in that list")
        else:
            if workbook['Member Name'].iloc[index] not in memberDict[workbook['Member Mail'].iloc[index].split('@')[1]]:
                memberDict[workbook['Member Mail'].iloc[index].split('@')[1]].append(workbook['Member Name'].iloc[index])
            else:
                print("Value is already in that list")

        if index == 7229:
            df = pd.DataFrame.from_dict(memberDict, orient='index')
            df.to_excel("Test.xlsx", sheet_name="Memberoutput")
            print("member list: " + f'{memberDict}')