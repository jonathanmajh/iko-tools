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
select * from 
(select workorder.wonum, workorder.description, longdescription.ldtext,
workorder.assetnum, asset.description as assetdesc, workorder.worktype,workorder.reportedby,
workorder.reportdate, workorder.wopriority, iko_conditions, workorder.status,
targstartdate, schedstart, actfinish, workorder.siteid,workorder.historyflag,
workorder.downtime, workorder.iko_downtime, iko_top3breakdown, workorder.failurecode,
iko_symptom, iko_desc1, iko_prodacreason
from workorder
left join longdescription on ldownertable = 'workorder' and workorder.workorderid = longdescription.ldkey
left join asset on asset.assetnum = workorder.assetnum and asset.siteid = workorder.siteid
where 
workorder.woclass = 'workorder'
and workorder.EXTERNALREFID is not null
and workorder.siteid = 'GS'
) t1
left join
(select itemnum, quantity, refwo, siteid from matusetrans) t2
on t1.wonum = t2.refwo and t1.siteid = t2.siteid
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
        if (isinstance(row[17], str)):
            row[17] = re.sub(CLEANR, '', row[17])
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