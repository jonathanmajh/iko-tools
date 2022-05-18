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

cursor.execute("""select t1.siteid, t1.assetnum as parentassetnum, t1.assetdescription as parentassetdescription, t1.pmnum as parentpmnumber, t1.jpnum as parentjobplan,
coalesce(t3.assetnum, t1.assetnum) as childasset, coalesce(t3.childassetdescription, t1.assetdescription) as childassetdescription,
case when ISNULL(t3.route,'bla') = 'bla' then 'No' else 'Yes' end as isroute, t3.route as routenumber, t3.jpnum as childjobplans, 
--bla is placeholder since the original has 0, but a string got coalesced and resulted in a string = 0 comparison error
t4.jptask as childtask, t4.description as childtaskshortdescription, t4.ldtext as childtasklongdescription
from 
(select pm.siteid, pm.assetnum, asset.description as assetdescription, pm.pmnum, pm.description as pmdescription, pm.jpnum, pm.route, pm.status from pm
left join asset on asset.assetnum = pm.assetnum and asset.siteid = pm.siteid
where pm.siteid = 'GH') t1
left join
(select route_stop.route, route_stop.assetnum, asset.description as childassetdescription, route_stop.jpnum, route_stop.stopsequence * 10 as taskid, route_stop.description from route_stop
left join asset on asset.assetnum = route_stop.assetnum and asset.siteid = 'pbm'
where route_stop.siteid = 'GH') t3
on t1.route = t3.route
left join
(select jobtask.jpnum, jobtask.siteid, jobtask.jptask, jobtask.description, longdescription.ldtext from jobtask 
left join longdescription on ldownertable = 'jobtask' and longdescription.ldkey = jobtask.jobtaskid
where jobtask.siteid = 'GH') t4
on coalesce(t3.jpnum, t1.jpnum) = t4.jpnum
order by parentpmnumber, childasset """)

wb = Workbook()
sheet = wb.create_sheet('Data')
sheet.append(
    ['siteid','parentassetnum','parentassetdescription','parentpmnumber','parentjobplan','childasset','childassetdescription','isroute','routenumber','childjobplans','childtask','childtaskshortdescription','childtasklongdescription']
)

CLEANR = re.compile('<.*?>') 

for row in cursor:
    try:
        row = list(row)
        if (isinstance(row[12], str)):
            row[12] = re.sub(CLEANR, '', row[12])
        sheet.append(row)
    except IllegalCharacterError:
        row[12] = row[12].replace('\x19', "'")
        row[12] = row[12].replace('\x13', "'")
        row[12] = row[12].replace('\x1d', "'")
        row[12] = row[12].replace('\x1c', "'")
        try:
            sheet.append(row)
        except Exception:
            print(row[12])

conn.close()

wb.save(save_path)