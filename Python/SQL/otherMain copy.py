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
-- overdue
select top 1 *, 'Overdue' as 'overdue/ontime/reject' from (
SELECT COALESCE(t2.assetnum, t1.assetnum) as assetnum, t1.siteid, pmnum, wonum, description, targcompdate, iko_overduedays, status, t1.route FROM (  
SELECT assetnum, pmnum,wonum,description, targcompdate, iko_overduedays, status, siteid, route FROM workorder  
WHERE (status='MISSED' OR iko_overduedays > 0)  
AND parent IS NULL  
AND istask=0  
--AND siteid='GV'  
AND targcompdate>'1/1/2022' 
AND targcompdate<'7/1/2022' 
AND STATUS <> 'REJECTED' 
) AS t1 
LEFT JOIN (  
SELECT route, assetnum, siteid FROM route_stop 
) AS t2  
ON t1.route=t2.route AND t1.siteid = t2.siteid) as t3
left join (
select assetnum, siteid, priority from asset) as t4
on t3.assetnum=t4.assetnum and t3.siteid=t4.siteid

UNION -- ontime

select top 1 *, 'Ontime' as 'overdue/ontime/reject' from (
SELECT COALESCE(t2.assetnum, t1.assetnum) as assetnum, t1.siteid, pmnum, wonum, description, targcompdate, iko_overduedays, status, t1.route FROM (  
SELECT assetnum, pmnum,wonum,description, targcompdate, iko_overduedays, status, siteid, route FROM workorder  
WHERE (iko_overduedays is null OR iko_overduedays = 0)  
AND parent IS NULL  
AND istask=0  
--AND siteid='GV'  
AND targcompdate>'1/1/2022' 
AND targcompdate<'7/1/2022' 
AND STATUS <> 'REJECTED' 
and status in ('CLOSE','FINISHED','COMP','WAITCLOSE') 
) AS t1 
LEFT JOIN (  
SELECT route, assetnum, siteid FROM route_stop 
) AS t2  
ON t1.route=t2.route AND t1.siteid = t2.siteid) as t3
left join (
select assetnum, siteid, priority from asset) as t4
on t3.assetnum=t4.assetnum and t3.siteid=t4.siteid

union --rejected
select top 1 *, 'Rejected' as 'overdue/ontime/reject' from (
SELECT COALESCE(t2.assetnum, t1.assetnum) as assetnum, t1.siteid, pmnum, wonum, description, targcompdate, iko_overduedays, status, t1.route FROM (  
SELECT assetnum, pmnum,wonum,description, targcompdate, iko_overduedays, status, siteid, route FROM workorder  
WHERE 
parent IS NULL  
AND istask=0  
--AND siteid='GV'  
AND targcompdate>'1/1/2022' 
AND targcompdate<'7/1/2022' 
AND STATUS = 'REJECTED' 
) AS t1 
LEFT JOIN (  
SELECT route, assetnum, siteid FROM route_stop 
) AS t2  
ON t1.route=t2.route AND t1.siteid = t2.siteid) as t3
left join (
select assetnum, siteid, priority from asset) as t4
on t3.assetnum=t4.assetnum and t3.siteid=t4.siteid
 """)

wb = Workbook()
sheet = wb.create_sheet('Data')
sheet.append(
    ['assetnum','siteid','pmnum','wonum','description','targcompdate','iko_overduedays','status','route','assetnum','siteid','priority','overdue/ontime/reject',]
)

for row in cursor:
    sheet.append(list(row))

conn.close()

wb.save(save_path)