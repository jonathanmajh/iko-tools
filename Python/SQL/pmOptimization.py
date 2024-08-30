from datetime import datetime
import re

import pyodbc
from openpyxl import Workbook
from openpyxl.utils.exceptions import IllegalCharacterError


conn = pyodbc.connect(
    "DRIVER={ODBC Driver 18 for SQL Server};SERVER=NSCANDACMAXSQL1\\MAXIMO;DATABASE=mxprod76;Trusted_Connection=yes;Encrypt=no;"
)

conn.setdecoding(pyodbc.SQL_CHAR, encoding="latin1")
conn.setdecoding(pyodbc.SQL_WCHAR, encoding="latin1")

save_path = f'C:\\Users\\majona\\Documents\\Code\\iko-tools\\Python\\SQL\\{datetime.now().strftime("%d-%m-%Y_%H-%M-%S")}.xlsx'

cursor = conn.cursor()

cursor.execute(
    """
SELECT eauditusername
    , eaudittimestamp
    , eaudittype
    , siteid
    , frequency
    , frequnit
    , pmnum
    , pmuid
FROM aws.maxprd.dbo.A_PM
WHERE eaudittype != 'D'
    AND pmuid IN (
        SELECT pmuid
        FROM aws.maxprd.dbo.a_pm
        WHERE eaudittype = 'u'
            AND datepart(year, eaudittimestamp) = 2024
        )
ORDER BY eaudittimestamp DESC
 """
)

wb = Workbook()
# get PM audit data
sheet = wb.create_sheet("pmData")
sheet.append(
    [
        "eauditusername",
        "eaudittimestamp",
        "eaudittype", #2
        "siteid",
        "frequency", #4
        "frequnit", #5
        "pmnum",
        "pmuid", #7
    ]
)

CLEANR = re.compile("<.*?>")

pms = {}
updatedPms = {}

for row in cursor:
    try:
        row = list(row)
        # if (isinstance(row[17], str)):
        #     row[17] = re.sub(CLEANR, '', row[17])
        sheet.append(row)
    except IllegalCharacterError:
        row[17] = row[17].replace("\x19", "'")
        row[17] = row[17].replace("\x13", "'")
        row[17] = row[17].replace("\x1d", "'")
        row[17] = row[17].replace("\x1c", "'")
        try:
            sheet.append(row)
        except Exception:
            print(row)
    if row[7] in pms.keys():
        if pms[row[7]] != [row[5], row[4]]:
            updatedPms[row[7]] = [row[3],row[6], row[1]]
    else: # add pm to dictionary
        pms[row[7]] = [row[5], row[4]]
        
sheet = wb.create_sheet("pmUpdated")
for pm in updatedPms.values():
    sheet.append(pm)


updatedJobplans = {}
# get job plan audit data
cursor.execute(
"""
SELECT eauditusername
    , eaudittimestamp
    , eaudittype
    , siteid
    , jobplanid
    , jpnum
    , orgid
FROM aws.maxprd.dbo.a_jobplan
WHERE eaudittype = 'U'
    AND datepart(year, eaudittimestamp) = 2024
ORDER BY eaudittimestamp
 """
)
sheet = wb.create_sheet("jobplanData")
sheet.append(
    [
        "eauditusername",
        "eaudittimestamp",
        "eaudittype",
        "siteid",
        "jobplanid", #4
        "jpnum",
        "orgid",
    ]
)

CLEANR = re.compile("<.*?>")

for row in cursor:
    try:
        row = list(row)
        # if (isinstance(row[17], str)):
        #     row[17] = re.sub(CLEANR, '', row[17])
        sheet.append(row)
    except IllegalCharacterError:
        row[17] = row[17].replace("\x19", "'")
        row[17] = row[17].replace("\x13", "'")
        row[17] = row[17].replace("\x1d", "'")
        row[17] = row[17].replace("\x1c", "'")
        try:
            sheet.append(row)
        except Exception:
            print(row)
    updatedJobplans[row[4]] = [row[3], row[5]]

# get jobtask audit data
cursor.execute(
    """
SELECT eauditusername
    , eaudittimestamp
    , eaudittype
    , siteid
    , orgid
    , jpnum
    , jobtaskid
FROM aws.maxprd.dbo.a_jobtask
WHERE eaudittype = 'U'
    AND datepart(year, eaudittimestamp) = 2024
ORDER BY eaudittimestamp
 """
)
sheet = wb.create_sheet("jobtasData")
sheet.append(
    [
        "eauditusername",
        "eaudittimestamp",
        "eaudittype",
        "siteid", #3
        "jobplanid",
        "jpnum", #5
        "orgid",
    ]
)

CLEANR = re.compile("<.*?>")

for row in cursor:
    try:
        row = list(row)
        # if (isinstance(row[17], str)):
        #     row[17] = re.sub(CLEANR, '', row[17])
        sheet.append(row)
    except IllegalCharacterError:
        row[17] = row[17].replace("\x19", "'")
        row[17] = row[17].replace("\x13", "'")
        row[17] = row[17].replace("\x1d", "'")
        row[17] = row[17].replace("\x1c", "'")
        try:
            sheet.append(row)
        except Exception:
            print(row)
    updatedJobplans[row[4]] = [row[3], row[5]]

sheet = wb.create_sheet("jobplanUpdated")
for pm in updatedJobplans.values():
    sheet.append(list(pm))

conn.close()

wb.save(save_path)
