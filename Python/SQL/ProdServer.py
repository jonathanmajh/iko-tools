from datetime import datetime
import re

import pyodbc
from openpyxl import Workbook
from openpyxl.utils.exceptions import IllegalCharacterError 


conn = pyodbc.connect('DRIVER={ODBC Driver 18 for SQL Server};SERVER=NSCANDACMAXSQL1\MAXIMO;DATABASE=mxprod76;Trusted_Connection=yes;Encrypt=no;')

conn.setdecoding(pyodbc.SQL_CHAR, encoding='latin1')
conn.setdecoding(pyodbc.SQL_WCHAR, encoding='latin1')

save_path = f'C:\\Users\\majona\\GitHub\\iko-tools\\Python\\SQL\\{datetime.now().strftime("%d-%m-%Y_%H-%M-%S")}.xlsx'

cursor = conn.cursor()

cursor.execute("""
select * from workorder 
left join longdescription on ldownertable = 'WORKLOG' and longdescription.ldkey = workorder.workorderid
where siteid = 'cam' and wonum like 'w23a%'
 """)

wb = Workbook()
sheet = wb.create_sheet('Data')
sheet.append(
    ['wonum','description','ldtext','assetnum','assetdesc','worktype','reportedby','reportdate','wopriority','iko_conditions','status','targstartdate','schedstart','actfinish','siteid','historyflag','downtime','iko_downtime','iko_top3breakdown','failurecode','iko_symptom','iko_desc1','iko_prodacreason','itemnum','quantity','refwo','siteid'
]
)

CLEANR = re.compile('<.*?>') 

for row in cursor:
    try:
        row = list(row)
        # if (isinstance(row[17], str)):
        #     row[17] = re.sub(CLEANR, '', row[17])
        sheet.append(row)
    except IllegalCharacterError:
        row[17] = row[17].replace('\x19', "'")
        row[17] = row[17].replace('\x13', "'")
        row[17] = row[17].replace('\x1d', "'")
        row[17] = row[17].replace('\x1c', "'")
        try:
            sheet.append(row)
        except Exception:
            print(row[17])

conn.close()

wb.save(save_path)