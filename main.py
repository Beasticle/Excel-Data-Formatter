from collections import defaultdict
import pandas as pd

workbook = pd.read_excel("EMEATeamsPermissions.xlsx", sheet_name="EMEATeamsPermissions")
workbook.head()

memberDict = defaultdict(list)
#domainNames = workbook['Member Mail']

for index, domainName in enumerate(workbook['Member Mail']):
    print(domainName + f" {index}")
    # count += 1
    if domainName == workbook['Member Mail'].iloc[index]:
        print(domainName + " " + workbook['Member Mail'].iloc[index].split('@')[1] + f"\n")
        #currentDN = workbook['Member Mail'].iloc[count]
        #print(memberDict)
        # memberDict.append(workbook['Member Name'].iloc[index] + " " + workbook['Member Mail'].iloc[index].split('@')[1])
        memberDict[workbook['Member Mail'].iloc[index].split('@')[1]].append(workbook['Member Name'].iloc[index])
        if index == 7229:
            df = pd.DataFrame.from_dict([memberDict])
            df.to_excel("Test.xlsx", sheet_name="Memberoutput")
            print("member list: " + f'{memberDict}')
            # break
        
    # else:
    #     print("how the fuck did we get here?")

