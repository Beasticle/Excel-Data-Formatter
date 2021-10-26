import pandas as pd

workbook = pd.read_excel("EMEATeamsPermissions.xlsx", sheet_name="EMEATeamsPermissions")
workbook.head()

count = 0
memberList = []
#domainNames = workbook['Member Mail']

for index, domainName in enumerate(workbook['Member Mail'], start=1):
    print(domainName + f" {count}")
    # count += 1
    if domainName == workbook['Member Mail'].iloc[index]:
        print(domainName + " " + workbook['Member Mail'].iloc[index].split('@')[1] + f"\n")
        #currentDN = workbook['Member Mail'].iloc[count]
        #print(memberDict)
        memberList.append(workbook['Member Name'].iloc[index] + " " + workbook['Member Mail'].iloc[index].split('@')[1])
        df = pd.DataFrame(memberList)
        df.to_excel("Test.xlsx", sheet_name="Memberoutput")
        
    # else:
    #     print("how the fuck did we get here?")

