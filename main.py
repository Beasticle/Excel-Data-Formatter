import pandas as pd

workbook = pd.read_excel("EMEATeamsPermissions.xlsx", sheet_name="EMEATeamsPermissions")
workbook.head()

#count = 2

#teamName = workbook['Teamname'].iloc(count)

memberList = []

for teamName in workbook['Teams Name']:
    count = 2
    count += 1
    print(teamName)
    if teamName == workbook['Teams Name'].iloc[count]:
        prevTeamName = workbook['Teams Name'].iloc[count]
        memberList.append(workbook['Member Name'].iloc[count])
        df = pd.DataFrame(memberList)
        df.to_excel("EMEATeamsPermissions.xlsx", sheet_name="Memberoutput")

