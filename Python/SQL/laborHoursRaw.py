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

sites = ['AA','BA','BL','CA','GC','GE','GH','GI','GK','GM','GP','GR','GS','GV']
dates = [[2024,8],[2024,7],[2024,6],[2024,5],[2024,4],[2024,3],[2024,2],[2024,1],[2023,12],[2023,11],[2023,10],[2023,9]]

sql =  """
SELECT kronos.laborcode
    , kronos.personname
    , kronos.kronoshours
    , maximo.maximohours
    , kronos.worksite
FROM (
    SELECT labor.worksite
        , attendance.laborcode
        , upper(p.lastname) + ', ' + upper(p.firstname) AS personname
        , sum(laborhours) kronoshours
    FROM aws.maxprd.dbo.attendance
    LEFT JOIN aws.maxprd.dbo.labor
        ON attendance.laborcode = labor.laborcode
    INNER JOIN aws.maxprd.dbo.person p
        ON attendance.laborcode = p.personid
    WHERE labor.worksite = 'XXXX'
        AND attendance.laborcode <> 'SALMGREG'
        AND attendance.include = 1
        AND DATEPART(month, startdate) = MMMM
        AND DATEPART(year, startdate) = YYYY
    GROUP BY labor.worksite
        , attendance.laborcode
        , p.lastname
        , p.firstname
    ) kronos
LEFT JOIN (
    SELECT w.siteid
        , al.laborcode
        , cast(isnull(sum(regularhrs), 0) + isnull(sum(premiumpayhours), 0) AS DECIMAL(12, 2)) maximohours
    FROM aws.maxprd.dbo.workorder w
    INNER JOIN aws.maxprd.dbo.labtrans al
        ON w.wonum = al.refwo
            AND w.siteid = al.siteid
    WHERE (
            al.vendor IS NULL
            OR al.laborcode = 'GPCONT01'
            )
        AND w.siteid = 'XXXX'
        AND al.laborcode <> 'SALMGREG'
        AND DATEPART(month, al.startdate) = MMMM
        AND DATEPART(year, al.startdate) = YYYY
        AND al.craft <> 'CORP'
    GROUP BY w.siteid
        , al.laborcode
    ) maximo
    ON kronos.worksite = maximo.siteid
        AND kronos.laborcode = maximo.laborcode
 """

cursor = conn.cursor()

wb = Workbook()
# get PM audit data
sheet = wb.create_sheet("Data")

sheet.append(
    [
        "laborcode",
        "personname",
        "kronoshours", #2
        "maximohours",
        "siteid", #4
        "year",
        "month",
    ]
)

for site in sites:
    for date in dates:
        print(f'{site} {date}')
        runSql = sql.replace('XXXX', site).replace('MMMM', str(date[1])).replace('YYYY',str(date[0]))
        cursor.execute(runSql)

        for row in cursor:
            row = list(row)
            row.append(date[0])
            row.append(date[1])
            sheet.append(row)

conn.close()

wb.save(save_path)
