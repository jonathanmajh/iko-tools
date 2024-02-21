from datetime import datetime
import re

import pyodbc
from openpyxl import Workbook
from openpyxl.utils.exceptions import IllegalCharacterError 


conn = pyodbc.connect('DRIVER={ODBC Driver 18 for SQL Server};SERVER=ESNLDKLUSQL1\GENERAL;DATABASE=PiranaCMMS;Trusted_Connection=yes;Encrypt=no;')

conn.setdecoding(pyodbc.SQL_CHAR, encoding='latin1')
conn.setdecoding(pyodbc.SQL_WCHAR, encoding='latin1')

save_path = f'C:\\Users\\majona\\Documents\\Code\\iko-tools\\Python\\SQL\\photos\\'

cursor = conn.cursor()

cursor.execute("""
select id, main_image from image_library
 """)

for row in cursor:
    image = row[1]
    filename = f'{row[0]}.jpg'
    with open(save_path + filename, 'wb') as new_jpg:
        new_jpg.write(image)

conn.close()