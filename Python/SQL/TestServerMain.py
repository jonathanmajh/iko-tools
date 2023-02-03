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
select t1.siteid, t1.iko_pmpackage, t1.assetnum as parentassetnum, t1.description as parentassetdesc, t1.pmnum as parentpmnum, t1.frequency, t1.frequnit,
t1.status, t1.jpnum, t2.assetnum as childassetnum, t2.description as childassetdesc, case when ISNULL(t2.route,'bla') = 'bla' then 'No' else 'Yes' end as isroute,
t2.route, t2.jpnum, t3.iko_conditions, t3.jptask, t3.description as taskdesc, t3.ldtext as longdescription, t3.metername from
(select pm.siteid, jobplan.iko_pmpackage, pm.assetnum, asset.description, pm.pmnum, pm.frequency, pm.frequnit, pm.status, pm.jpnum, pm.route from pm
left join jobplan on pm.jpnum = jobplan.jpnum and pm.siteid = jobplan.siteid
left join asset on pm.assetnum = asset.assetnum and pm.siteid = asset.siteid) t1
left join
(select route, route_stop.assetnum, asset.description, jpnum, route_stop.siteid, routestopid from route_stop
left join asset on route_stop.assetnum = asset.assetnum and route_stop.siteid = asset.siteid) t2
on t1.siteid = t2.siteid and t1.route = t2.route
left join
(select jobtask.jpnum, jobtask.jptask, jobtask.description, jobtask.metername, jobtask.siteid, jobplan.IKO_CONDITIONS, longdescription.ldtext from jobtask
left join jobplan on jobplan.jpnum = jobtask.jpnum and jobplan.siteid = jobtask.siteid
left join longdescription on ldownertable = 'jobtask' and longdescription.ldkey = jobtask.jobtaskid) t3
on t1.siteid = t3.siteid and t3.jpnum = coalesce(t2.jpnum, t1.jpnum)
where t1.siteid <> 'pbm'
order by t1.siteid, t1.pmnum, t2.assetnum """)

wb = Workbook()
sheet = wb.create_sheet('Data')
sheet.append(
    ['siteid','parentassetnum','parentassetdescription','parentpmnumber','parentjobplan','childasset','childassetdescription','isroute','routenumber','childjobplans','childtask','childtaskshortdescription','childtasklongdescription']
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