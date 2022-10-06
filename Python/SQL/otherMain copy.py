from datetime import datetime
import re

import pyodbc
from openpyxl import Workbook
from openpyxl.utils.exceptions import IllegalCharacterError 


conn = pyodbc.connect('DRIVER={ODBC Driver 18 for SQL Server};SERVER=NSMAXIM1SQL01\MAXIMO;DATABASE=mxtest76;UID=maximo;PWD=M@XIMO1;Encrypt=no')

conn.setdecoding(pyodbc.SQL_CHAR, encoding='latin1')
conn.setdecoding(pyodbc.SQL_WCHAR, encoding='latin1')

save_path = f'C:\\Users\\majona\\GitHub\\iko-tools\\Python\\SQL\\{datetime.now().strftime("%d-%m-%Y_%H-%M-%S")}.xlsx'

cursor = conn.cursor()

cursor.execute("""
select orgid, siteid, jpnum, jptask, metername, description, ldtext from jobtask
left join longdescription on ldownertable = 'jobtask' and jobtask.jobtaskid = longdescription.ldkey
where metername is not null
 """)

wb = Workbook()
sheet = wb.create_sheet('Data')
sheet.append(
    ['orgid','siteid','jpnum','jptask','metername','description','details']
)

for row in cursor:
    sheet.append(list(row))

conn.close()

wb.save(save_path)