from datetime import datetime
import re

import pyodbc
from openpyxl import Workbook
from openpyxl.utils.exceptions import IllegalCharacterError 


conn = pyodbc.connect('DRIVER={ODBC Driver 18 for SQL Server};SERVER=ESNLDKLUPIRAN1\SQLEXPRESS;DATABASE=PiranaCMMS;Trusted_Connection=yes;Encrypt=no;')

conn.setdecoding(pyodbc.SQL_CHAR, encoding='latin1')
conn.setdecoding(pyodbc.SQL_WCHAR, encoding='latin1')

save_path = f'C:\\Users\\majona\\GitHub\\iko-tools\\Python\\SQL\\{datetime.now().strftime("%d-%m-%Y_%H-%M-%S")}.xlsx'

cursor = conn.cursor()

cursor.execute("""
select * from Personnel
 """)

wb = Workbook()
sheet = wb.create_sheet('Data')
sheet.append(
    ['orgid','siteid','jpnum','jptask','metername','description','details']
)

for row in cursor:
    temp = [str(x).encode('ascii',errors='ignore') for x in list(row)]
    sheet.append(temp)

conn.close()

wb.save(save_path)