#Automate OEE data reports for Jonathan
#Ishmam Raza Dewan
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import datetime, time
import xlwings as xw
import win32com.client as win32
from pathlib import Path
import pandas as pd
import re
import os


########### FUNCTION DEFINITIONS

def wait(a,b): #a is original file name, b is new file name
    try:
        os.rename(a, b)
    except:
        q=0

    while q == 0: #there is probably a smarter way of doing this lol
        if q==0:
            time.sleep(3)
            try:
                os.rename(a, b)
            except:
                q=0
            else:
                q=1

###############################


browser = webdriver.Edge() #ensure your browser version and web driver version match

browser.get('http://nsikobrmiis1n/PRODAC/Reports/ProdacReportsMain.aspx') #Site AA
browser.maximize_window()
time.sleep(3)

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

GoTo = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/div/div[2]/ul/li/ul/li[1]/ul/li[5]/a")
GoTo.click()

time.sleep(3)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

FromDate = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/input")
FromDate.click()
FromDate.send_keys('01/01/2023')

GroupingDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/table/tbody/tr/td[2]/img")
GroupingDropDown.click()
time.sleep(1)
GroupingMonth = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/div/div/div/ul/li[2]")
GroupingMonth.click()
time.sleep(1)
ReportPeriodDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/table/tbody/tr/td[2]/img")
ReportPeriodDropDown.click()

time.sleep(1)

ReportYTD = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/div/div/div/ul/li[4]")
ReportYTD.click()

time.sleep(1)

PartGroupsDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/table/tbody/tr/td[2]/img")
PartGroupsDropDown.click()
time.sleep(1)
NoPartGroups = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/div/div/div/ul/li[1]")
NoPartGroups.click()

time.sleep(1)

ShowInDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[2]/img")
ShowInDropDown.click()
time.sleep(1)
ShowInExcel = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/div/div/div/ul/li[3]")
ShowInExcel.click()


iframeSL = browser.find_element(By.ID, "frmGrids")
browser.switch_to.frame(iframeSL)

SL = browser.find_element(By.XPATH, "/html/body/form/div[3]/table/tbody/tr[1]/td[1]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/div/table/tbody/tr[1]/td[1]/table/tbody[2]/tr/td/div[2]/table/tbody/tr/td[4]")
SL.click()

browser.switch_to.default_content()

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

Generate = browser.find_element(By.ID, "btnGenerate__4")
Generate.click()

time.sleep(5)
wait('IkoManagementKPI_Summary.xls', 'AA.xls')

browser.quit()


############################# AA - EXCEL SECTION ###############################


i = 2
q = 2
while i/q < 40:
    wb = xw.Book('AA.xls').sheets[0]
    
    i = i + 1
    if wb['B' + str(i)].value == 'OEE (%)':
        LINE = wb['B' + str(i-9)].value
        print(LINE)
        OEEs = wb.range('C' + str(i) + ':Q' + str(i)).value
        q=q+1
        wb = xw.Book('OEE Summary.xlsx').sheets[0]
        wb['B' + str(q)].value = LINE
        wb['A' + str(q)].value = 'AA'
        wb.range('C' + str(q) + ':Q' + str(q)).value = OEEs
        wb = xw.Book('OEE Summary.xlsx')
        wb.save()

wb = xw.Book('AA.xls')
wb.app.quit()

############################ GK - WEBSITE SECTION ###############################
browser = webdriver.Edge()
browser.get('http://nsigbrmiis1/PRODAC/Reports/ProdacReportsMain.aspx')
browser.maximize_window()

time.sleep(3)

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

GoTo = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/div/div[2]/ul/li/ul/li[1]/ul/li[5]/a")
GoTo.click()

time.sleep(3)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

FromDate = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/input")
FromDate.click()
FromDate.send_keys('01/01/2023')

GroupingDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/table/tbody/tr/td[2]/img")
GroupingDropDown.click()
time.sleep(1)
GroupingMonth = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/div/div/div/ul/li[2]")
GroupingMonth.click()
time.sleep(1)
ReportPeriodDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/table/tbody/tr/td[2]/img")
ReportPeriodDropDown.click()

time.sleep(1)

ReportYTD = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/div/div/div/ul/li[4]")
ReportYTD.click()

time.sleep(1)

PartGroupsDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/table/tbody/tr/td[2]/img")
PartGroupsDropDown.click()
time.sleep(1)
NoPartGroups = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/div/div/div/ul/li[1]")
NoPartGroups.click()

time.sleep(1)

ShowInDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[2]/img")
ShowInDropDown.click()
time.sleep(1)
ShowInExcel = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/div/div/div/ul/li[3]")
ShowInExcel.click()

iframeSL = browser.find_element(By.ID, "frmGrids")
browser.switch_to.frame(iframeSL)

RL = browser.find_element(By.XPATH, "/html/body/form/div[3]/table/tbody/tr[1]/td[1]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/div/table/tbody/tr[1]/td[1]/table/tbody[2]/tr/td/div[2]/table/tbody/tr/td[4]")
RL.click()
time.sleep(2)

browser.switch_to.default_content()

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

Generate = browser.find_element(By.ID, "btnGenerate__4")
Generate.click()

time.sleep(5)
wait('IkoManagementKPI_RL_NA_Summary.xls','GK.xls')
browser.quit()

############################ GK - EXCEL SECTION ###############################

i = 2
while i/q < 90:
    wb = xw.Book('GK.xls').sheets[0]
    
    i = i + 1
    if wb['B' + str(i)].value == 'OEE (%)':
        LINE = wb['B' + str(i-10)].value
        OEEs = wb.range('C' + str(i) + ':Q' + str(i)).value
        q=q+1
        wb = xw.Book('OEE Summary.xlsx').sheets[0]
        wb['B' + str(q)].value = LINE
        wb['A' + str(q)].value = 'GK'
        wb.range('C' + str(q) + ':Q' + str(q)).value = OEEs
        wb = xw.Book('OEE Summary.xlsx')
        wb.save()


wb = xw.Book('GK.xls')
wb.app.quit()

############################ GKB6 - WEBSITE SECTION ###############################

browser = webdriver.Edge()
browser.get('http://nsigbrmiis1/PRODAC/Reports/ProdacReportsMain.aspx')
browser.maximize_window()

time.sleep(3)

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

GoTo = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/div/div[2]/ul/li/ul/li[1]/ul/li[7]/a")
GoTo.click()

time.sleep(3)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

FromDate = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/input")
FromDate.click()
FromDate.send_keys('01/01/2023')

GroupingDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/table/tbody/tr/td[2]/img")
GroupingDropDown.click()
time.sleep(1)
GroupingMonth = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/div/div/div/ul/li[2]")
GroupingMonth.click()
time.sleep(1)
ReportPeriodDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/table/tbody/tr/td[2]/img")
ReportPeriodDropDown.click()

time.sleep(1)

ReportYTD = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/div/div/div/ul/li[4]")
ReportYTD.click()

time.sleep(1)

PartGroupsDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/table/tbody/tr/td[2]/img")
PartGroupsDropDown.click()
time.sleep(1)
NoPartGroups = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/div/div/div/ul/li[1]")
NoPartGroups.click()

time.sleep(1)

ShowInDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[2]/img")
ShowInDropDown.click()
time.sleep(1)
ShowInExcel = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/div/div/div/ul/li[3]")
ShowInExcel.click()

iframeSL = browser.find_element(By.ID, "frmGrids")
browser.switch_to.frame(iframeSL)

B6 = browser.find_element(By.XPATH, "/html/body/form/div[3]/table/tbody/tr[1]/td[1]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/div/table/tbody/tr[1]/td[1]/table/tbody[2]/tr/td/div[2]/table/tbody/tr/td[4]")
B6.click()
time.sleep(2)

browser.switch_to.default_content()

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

Generate = browser.find_element(By.ID, "btnGenerate__4")
Generate.click()

time.sleep(5)
wait('IkoManagementKPI_RL_PROTECTOBOARD_Summary.xls','GKB6.xls')
browser.quit()

############################ GKB6 - EXCEL SECTION ###############################

i = 2
while i/q < 90:
    wb = xw.Book('GKB6.xls').sheets[0]
    
    i = i + 1
    if wb['B' + str(i)].value == 'OEE (%)':
        LINE = wb['B' + str(i-10)].value
        OEEs = wb.range('C' + str(i) + ':Q' + str(i)).value
        q=q+1
        wb = xw.Book('OEE Summary.xlsx').sheets[0]
        wb['B' + str(q)].value = LINE
        wb['A' + str(q)].value = 'GKB6'
        wb.range('C' + str(q) + ':Q' + str(q)).value = OEEs
        wb = xw.Book('OEE Summary.xlsx')
        wb.save()


wb = xw.Book('GKB6.xls')
wb.app.quit()

############################ GJ - WEBSITE SECTION ###############################


browser = webdriver.Edge()
browser.get('http://crctoriis1/ProdacNewWeb/Reports/ProdacReportsMain.aspx')
browser.maximize_window()

time.sleep(3)

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

GoTo = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/div/div[2]/ul/li/ul/li[1]/ul/li[5]/a")
GoTo.click()

time.sleep(3)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

FromDate = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/input")
FromDate.click()
FromDate.send_keys('01/01/2023')

GroupingDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/table/tbody/tr/td[2]/img")
GroupingDropDown.click()
time.sleep(1)
GroupingMonth = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/div/div/div/ul/li[2]")
GroupingMonth.click()
time.sleep(1)
ReportPeriodDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/table/tbody/tr/td[2]/img")
ReportPeriodDropDown.click()

time.sleep(1)

ReportYTD = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/div/div/div/ul/li[4]")
ReportYTD.click()

time.sleep(1)

PartGroupsDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/table/tbody/tr/td[2]/img")
PartGroupsDropDown.click()
time.sleep(1)
NoPartGroups = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/div/div/div/ul/li[1]")
NoPartGroups.click()

time.sleep(1)

ShowInDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[2]/img")
ShowInDropDown.click()
time.sleep(1)
ShowInExcel = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/div/div/div/ul/li[3]")
ShowInExcel.click()

iframeSL = browser.find_element(By.ID, "frmGrids")
browser.switch_to.frame(iframeSL)

SL = browser.find_element(By.XPATH, "/html/body/form/div[3]/table/tbody/tr[1]/td[1]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/div/table/tbody/tr[1]/td[1]/table/tbody[2]/tr/td/div[2]/table/tbody/tr/td[4]")
SL.click()
time.sleep(2)

browser.switch_to.default_content()

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

Generate = browser.find_element(By.ID, "btnGenerate__4")
Generate.click()

time.sleep(5)
wait('IkoManagementKPI_Summary.xls','GJ.xls')
browser.quit()

############################ GJ - EXCEL SECTION ###############################

i = 2
while i/q < 90:
    wb = xw.Book('GJ.xls').sheets[0]
    
    i = i + 1
    if wb['B' + str(i)].value == 'OEE (%)':
        LINE = wb['B' + str(i-9)].value
        OEEs = wb.range('C' + str(i) + ':Q' + str(i)).value
        q=q+1
        wb = xw.Book('OEE Summary.xlsx').sheets[0]
        wb['B' + str(q)].value = LINE
        wb['A' + str(q)].value = 'GJ'
        wb.range('C' + str(q) + ':Q' + str(q)).value = OEEs
        wb = xw.Book('OEE Summary.xlsx')
        wb.save()


wb = xw.Book('GJ.xls')
wb.app.quit()

############################ GH - WEBSITE SECTION ###############################


browser = webdriver.Edge()
browser.get('http://ikohawkiis2/ProdacNewWeb/Reports/ProdacReportsMain.aspx')
browser.maximize_window()

time.sleep(3)

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

GoTo = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/div/div[2]/ul/li/ul/li[1]/ul/li[6]/a")
GoTo.click()

time.sleep(3)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

FromDate = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/input")
FromDate.click()
FromDate.send_keys('01/01/2023')

GroupingDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/table/tbody/tr/td[2]/img")
GroupingDropDown.click()
time.sleep(1)
GroupingMonth = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/div/div/div/ul/li[2]")
GroupingMonth.click()
time.sleep(1)
ReportPeriodDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/table/tbody/tr/td[2]/img")
ReportPeriodDropDown.click()

time.sleep(1)

ReportYTD = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/div/div/div/ul/li[4]")
ReportYTD.click()

time.sleep(1)

PartGroupsDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/table/tbody/tr/td[2]/img")
PartGroupsDropDown.click()
time.sleep(1)
NoPartGroups = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/div/div/div/ul/li[1]")
NoPartGroups.click()

time.sleep(1)

ShowInDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[2]/img")
ShowInDropDown.click()
time.sleep(1)
ShowInExcel = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/div/div/div/ul/li[3]")
ShowInExcel.click()

iframeSL = browser.find_element(By.ID, "frmGrids")
browser.switch_to.frame(iframeSL)

SM = browser.find_element(By.XPATH, "/html/body/form/div[3]/table/tbody/tr[1]/td[1]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/div/table/tbody/tr[1]/td[1]/table/tbody[2]/tr/td/div[2]/table/tbody/tr/td[4]")
SM.click()

time.sleep(2)

browser.switch_to.default_content()

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

Generate = browser.find_element(By.ID, "btnGenerate__4")
Generate.click()

time.sleep(5)
wait('IkoManagementKPI_Summary.xls','GH.xls')
browser.quit()

############################ GH - EXCEL SECTION ###############################

i = 2
while i/q < 90:
    wb = xw.Book('GH.xls').sheets[0]
    
    i = i + 1
    if wb['B' + str(i)].value == 'OEE (%)':
        LINE = wb['B' + str(i-9)].value
        OEEs = wb.range('C' + str(i) + ':Q' + str(i)).value
        q=q+1
        wb = xw.Book('OEE Summary.xlsx').sheets[0]
        wb['B' + str(q)].value = LINE
        wb['A' + str(q)].value = 'GH'
        wb.range('C' + str(q) + ':Q' + str(q)).value = OEEs
        wb = xw.Book('OEE Summary.xlsx')
        wb.save()


wb = xw.Book('GH.xls')
wb.app.quit()

############################ BA - WEBSITE SECTION ###############################


browser = webdriver.Edge()
browser.get('http://ikocalgiis1/ProdacNewWeb/Reports/ProdacReportsMain.aspx')
browser.maximize_window()

time.sleep(3)

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

GoTo = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/div/div[2]/ul/li/ul/li[1]/ul/li[5]/a")
GoTo.click()

time.sleep(3)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

FromDate = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/input")
FromDate.click()
FromDate.send_keys('01/01/2023')

GroupingDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/table/tbody/tr/td[2]/img")
GroupingDropDown.click()
time.sleep(1)
GroupingMonth = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/div/div/div/ul/li[2]")
GroupingMonth.click()
time.sleep(1)
ReportPeriodDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/table/tbody/tr/td[2]/img")
ReportPeriodDropDown.click()

time.sleep(1)

ReportYTD = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/div/div/div/ul/li[4]")
ReportYTD.click()

time.sleep(1)

PartGroupsDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/table/tbody/tr/td[2]/img")
PartGroupsDropDown.click()
time.sleep(1)
NoPartGroups = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/div/div/div/ul/li[1]")
NoPartGroups.click()

time.sleep(1)

ShowInDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[2]/img")
ShowInDropDown.click()
time.sleep(1)
ShowInExcel = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/div/div/div/ul/li[3]")
ShowInExcel.click()

iframeSL = browser.find_element(By.ID, "frmGrids")
browser.switch_to.frame(iframeSL)

SL = browser.find_element(By.XPATH, "/html/body/form/div[3]/table/tbody/tr[1]/td[1]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/div/table/tbody/tr[1]/td[1]/table/tbody[2]/tr/td/div[2]/table/tbody/tr/td[4]")
SL.click()
time.sleep(2)

browser.switch_to.default_content()

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

Generate = browser.find_element(By.ID, "btnGenerate__4")
Generate.click()

time.sleep(5)
wait('IkoManagementKPI_Summary.xls','BA.xls')
browser.quit()

############################ BA - EXCEL SECTION ###############################

i = 2
while i/q < 90:
    wb = xw.Book('BA.xls').sheets[0]
    
    i = i + 1
    if wb['B' + str(i)].value == 'OEE (%)':
        LINE = wb['B' + str(i-9)].value
        OEEs = wb.range('C' + str(i) + ':Q' + str(i)).value
        q=q+1
        wb = xw.Book('OEE Summary.xlsx').sheets[0]
        wb['B' + str(q)].value = LINE
        wb['A' + str(q)].value = 'BA'
        wb.range('C' + str(q) + ':Q' + str(q)).value = OEEs
        wb = xw.Book('OEE Summary.xlsx')
        wb.save()


wb = xw.Book('BA.xls')
wb.app.quit()

############################ GC - WEBSITE SECTION ###############################


browser = webdriver.Edge()
browser.get('http://ikosumiis1/ProdacNewWeb/Reports/ProdacReportsMain.aspx')
browser.maximize_window()

time.sleep(3)

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

GoTo = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/div/div[2]/ul/li/ul/li[1]/ul/li[9]/a")
GoTo.click()

time.sleep(3)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

FromDate = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/input")
FromDate.click()
FromDate.send_keys('01/01/2023')

GroupingDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/table/tbody/tr/td[2]/img")
GroupingDropDown.click()
time.sleep(1)
GroupingMonth = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/div/div/div/ul/li[2]")
GroupingMonth.click()
time.sleep(1)
ReportPeriodDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/table/tbody/tr/td[2]/img")
ReportPeriodDropDown.click()

time.sleep(1)

ReportYTD = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/div/div/div/ul/li[4]")
ReportYTD.click()

time.sleep(1)

PartGroupsDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/table/tbody/tr/td[2]/img")
PartGroupsDropDown.click()
time.sleep(1)
NoPartGroups = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/div/div/div/ul/li[1]")
NoPartGroups.click()

time.sleep(1)

ShowInDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[2]/img")
ShowInDropDown.click()
time.sleep(1)
ShowInExcel = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/div/div/div/ul/li[3]")
ShowInExcel.click()

iframeSL = browser.find_element(By.ID, "frmGrids")
browser.switch_to.frame(iframeSL)

SL = browser.find_element(By.XPATH, "/html/body/form/div[3]/table/tbody/tr[1]/td[1]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/div/table/tbody/tr[1]/td[1]/table/tbody[2]/tr/td/div[2]/table/tbody/tr/td[4]")
SL.click()
time.sleep(2)

browser.switch_to.default_content()

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

Generate = browser.find_element(By.ID, "btnGenerate__4")
Generate.click()

time.sleep(5)
wait('IkoManagementKPI_Summary.xls','GC.xls')
browser.quit()

############################ GC - EXCEL SECTION ###############################

i = 2
while i/q < 90:
    wb = xw.Book('GC.xls').sheets[0]
    
    i = i + 1
    if wb['B' + str(i)].value == 'OEE (%)':
        LINE = wb['B' + str(i-9)].value
        OEEs = wb.range('C' + str(i) + ':Q' + str(i)).value
        q=q+1
        wb = xw.Book('OEE Summary.xlsx').sheets[0]
        wb['B' + str(q)].value = LINE
        wb['A' + str(q)].value = 'GC'
        wb.range('C' + str(q) + ':Q' + str(q)).value = OEEs
        wb = xw.Book('OEE Summary.xlsx')
        wb.save()


wb = xw.Book('GC.xls')
wb.app.quit()

############################ GCRL - WEBSITE SECTION ###############################


browser = webdriver.Edge()
browser.get('http://ikosumiis1/ProdacNewWeb/Reports/ProdacReportsMain.aspx')
browser.maximize_window()

time.sleep(3)

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

GoTo = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/div/div[2]/ul/li/ul/li[1]/ul/li[10]/a")
GoTo.click()

time.sleep(3)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

FromDate = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/input")
FromDate.click()
FromDate.send_keys('01/01/2023')

GroupingDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/table/tbody/tr/td[2]/img")
GroupingDropDown.click()
time.sleep(1)
GroupingMonth = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/div/div/div/ul/li[2]")
GroupingMonth.click()
time.sleep(1)
ReportPeriodDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/table/tbody/tr/td[2]/img")
ReportPeriodDropDown.click()

time.sleep(1)

ReportYTD = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/div/div/div/ul/li[4]")
ReportYTD.click()

time.sleep(1)

PartGroupsDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/table/tbody/tr/td[2]/img")
PartGroupsDropDown.click()
time.sleep(1)
NoPartGroups = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/div/div/div/ul/li[1]")
NoPartGroups.click()

time.sleep(1)

ShowInDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[2]/img")
ShowInDropDown.click()
time.sleep(1)
ShowInExcel = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/div/div/div/ul/li[3]")
ShowInExcel.click()

iframeSL = browser.find_element(By.ID, "frmGrids")
browser.switch_to.frame(iframeSL)

RL = browser.find_element(By.XPATH, "/html/body/form/div[3]/table/tbody/tr[1]/td[1]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/div/table/tbody/tr[1]/td[1]/table/tbody[2]/tr/td/div[2]/table/tbody/tr/td[4]")
RL.click()
time.sleep(2)

browser.switch_to.default_content()

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

Generate = browser.find_element(By.ID, "btnGenerate__4")
Generate.click()

time.sleep(5)
wait('IkoManagementKPI_RL_NA_Summary.xls','GCRL.xls')
browser.quit()

############################ GCRL - EXCEL SECTION ###############################

i = 2
while i/q < 90:
    wb = xw.Book('GCRL.xls').sheets[0]
    
    i = i + 1
    if wb['B' + str(i)].value == 'OEE (%)':
        LINE = wb['B' + str(i-10)].value
        OEEs = wb.range('C' + str(i) + ':Q' + str(i)).value
        q=q+1
        wb = xw.Book('OEE Summary.xlsx').sheets[0]
        wb['B' + str(q)].value = LINE
        wb['A' + str(q)].value = 'GCRL'
        wb.range('C' + str(q) + ':Q' + str(q)).value = OEEs
        wb = xw.Book('OEE Summary.xlsx')
        wb.save()


wb = xw.Book('GCRL.xls')
wb.app.quit()

############################ CA - WEBSITE SECTION ###############################


browser = webdriver.Edge()
browser.get('http://ikokankiis1/ProdacNewWeb/Reports/ProdacReportsMain.aspx')
browser.maximize_window()

time.sleep(3)

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

GoTo = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/div/div[2]/ul/li/ul/li[1]/ul/li[5]/a")
GoTo.click()

time.sleep(3)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

FromDate = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/input")
FromDate.click()
FromDate.send_keys('01/01/2023')

GroupingDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/table/tbody/tr/td[2]/img")
GroupingDropDown.click()
time.sleep(1)
GroupingMonth = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/div/div/div/ul/li[2]")
GroupingMonth.click()
time.sleep(1)
ReportPeriodDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/table/tbody/tr/td[2]/img")
ReportPeriodDropDown.click()

time.sleep(1)

ReportYTD = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/div/div/div/ul/li[4]")
ReportYTD.click()

time.sleep(1)

PartGroupsDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/table/tbody/tr/td[2]/img")
PartGroupsDropDown.click()
time.sleep(1)
NoPartGroups = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/div/div/div/ul/li[1]")
NoPartGroups.click()

time.sleep(1)

ShowInDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[2]/img")
ShowInDropDown.click()
time.sleep(1)
ShowInExcel = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/div/div/div/ul/li[3]")
ShowInExcel.click()

iframeSL = browser.find_element(By.ID, "frmGrids")
browser.switch_to.frame(iframeSL)

SL = browser.find_element(By.ID, "grdDepartment_chk")
SL.click()
time.sleep(2)

browser.switch_to.default_content()

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

Generate = browser.find_element(By.ID, "btnGenerate__4")
Generate.click()

time.sleep(5)
wait('IkoManagementKPI_Summary.xls','CA.xls')
browser.quit()

############################ CA - EXCEL SECTION ###############################

i = 2
while i/q < 90:
    wb = xw.Book('CA.xls').sheets[0]
    
    i = i + 1
    if wb['B' + str(i)].value == 'OEE (%)':
        LINE = wb['B' + str(i-9)].value
        OEEs = wb.range('C' + str(i) + ':Q' + str(i)).value
        q=q+1
        wb = xw.Book('OEE Summary.xlsx').sheets[0]
        wb['B' + str(q)].value = LINE
        wb['A' + str(q)].value = 'CA'
        wb.range('C' + str(q) + ':Q' + str(q)).value = OEEs
        wb = xw.Book('OEE Summary.xlsx')
        wb.save()


wb = xw.Book('CA.xls')
wb.app.quit()

############################ CARL - WEBSITE SECTION ###############################

browser = webdriver.Edge()
browser.get('http://ikokankiis1/ProdacNewWeb/Reports/ProdacReportsMain.aspx')
browser.maximize_window()

time.sleep(3)

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

GoTo = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/div/div[2]/ul/li/ul/li[1]/ul/li[10]/a")
GoTo.click()

time.sleep(3)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

FromDate = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/input")
FromDate.click()
FromDate.send_keys('01/01/2023')

GroupingDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/table/tbody/tr/td[2]/img")
GroupingDropDown.click()
time.sleep(1)
GroupingMonth = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/div/div/div/ul/li[2]")
GroupingMonth.click()
time.sleep(1)
ReportPeriodDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/table/tbody/tr/td[2]/img")
ReportPeriodDropDown.click()

time.sleep(1)

ReportYTD = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/div/div/div/ul/li[4]")
ReportYTD.click()

time.sleep(1)

PartGroupsDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/table/tbody/tr/td[2]/img")
PartGroupsDropDown.click()
time.sleep(1)
NoPartGroups = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/div/div/div/ul/li[1]")
NoPartGroups.click()

time.sleep(1)

ShowInDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[2]/img")
ShowInDropDown.click()
time.sleep(1)
ShowInExcel = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/div/div/div/ul/li[3]")
ShowInExcel.click()

iframeSL = browser.find_element(By.ID, "frmGrids")
browser.switch_to.frame(iframeSL)

RL = browser.find_element(By.XPATH, "/html/body/form/div[3]/table/tbody/tr[1]/td[1]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/div/table/tbody/tr[1]/td[1]/table/tbody[2]/tr/td/div[2]/table/tbody/tr/td[4]")
RL.click()
time.sleep(2)

browser.switch_to.default_content()

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

Generate = browser.find_element(By.ID, "btnGenerate__4")
Generate.click()

time.sleep(5)
wait('IkoManagementKPI_RL_NA_Summary.xls','CARL.xls')
browser.quit()


############################ CARL - EXCEL SECTION ###############################

i = 2
while i/q < 90:
    wb = xw.Book('CARL.xls').sheets[0]
    
    i = i + 1
    if wb['B' + str(i)].value == 'OEE (%)':
        LINE = wb['B' + str(i-10)].value
        OEEs = wb.range('C' + str(i) + ':Q' + str(i)).value
        q=q+1
        wb = xw.Book('OEE Summary.xlsx').sheets[0]
        wb['B' + str(q)].value = LINE
        wb['A' + str(q)].value = 'CARL'
        wb.range('C' + str(q) + ':Q' + str(q)).value = OEEs
        wb = xw.Book('OEE Summary.xlsx')
        wb.save()


wb = xw.Book('CARL.xls')
wb.app.quit()


############################ GS - WEBSITE SECTION ###############################


browser = webdriver.Edge()
browser.get('http://nsikosyliis1/ProdacNewWeb/Reports/ProdacReportsMain.aspx')
browser.maximize_window()

time.sleep(3)

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

GoTo = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/div/div[2]/ul/li/ul/li[1]/ul/li[5]/a")
GoTo.click()

time.sleep(3)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

FromDate = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/input")
FromDate.click()
FromDate.send_keys('01/01/2023')

GroupingDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/table/tbody/tr/td[2]/img")
GroupingDropDown.click()
time.sleep(1)
GroupingMonth = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/div/div/div/ul/li[2]")
GroupingMonth.click()
time.sleep(1)
ReportPeriodDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/table/tbody/tr/td[2]/img")
ReportPeriodDropDown.click()

time.sleep(1)

ReportYTD = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/div/div/div/ul/li[4]")
ReportYTD.click()

time.sleep(1)

PartGroupsDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/table/tbody/tr/td[2]/img")
PartGroupsDropDown.click()
time.sleep(1)
NoPartGroups = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/div/div/div/ul/li[1]")
NoPartGroups.click()

time.sleep(1)

ShowInDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[2]/img")
ShowInDropDown.click()
time.sleep(1)
ShowInExcel = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/div/div/div/ul/li[3]")
ShowInExcel.click()

iframeSL = browser.find_element(By.ID, "frmGrids")
browser.switch_to.frame(iframeSL)

All = browser.find_element(By.ID, "grdDepartment_chk")
All.click()
time.sleep(2)

browser.switch_to.default_content()

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

Generate = browser.find_element(By.ID, "btnGenerate__4")
Generate.click()

time.sleep(5)
wait('IkoManagementKPI_Summary.xls','GS.xls')
browser.quit()

############################ GS - EXCEL SECTION ###############################

i = 2
while i/q < 90:
    wb = xw.Book('GS.xls').sheets[0]
    
    i = i + 1
    if wb['B' + str(i)].value == 'OEE (%)':
        LINE = wb['B' + str(i-9)].value
        OEEs = wb.range('C' + str(i) + ':Q' + str(i)).value
        q=q+1
        wb = xw.Book('OEE Summary.xlsx').sheets[0]
        wb['B' + str(q)].value = LINE
        wb['A' + str(q)].value = 'GS'
        wb.range('C' + str(q) + ':Q' + str(q)).value = OEEs
        wb = xw.Book('OEE Summary.xlsx')
        wb.save()


wb = xw.Book('GS.xls')
wb.app.quit()

############################ GV - WEBSITE SECTION ###############################


browser = webdriver.Edge()
browser.get('http://nsikohiliis1.na.iko/PRODAC/Reports/ProdacReportsMain.aspx')
browser.maximize_window()

time.sleep(3)

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

GoTo = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/div/div[2]/ul/li/ul/li[1]/ul/li[5]/a")
GoTo.click()

time.sleep(3)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

FromDate = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/input")
FromDate.click()
FromDate.send_keys('01/01/2023')

GroupingDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/table/tbody/tr/td[2]/img")
GroupingDropDown.click()
time.sleep(1)
GroupingMonth = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/div/div/div/ul/li[2]")
GroupingMonth.click()
time.sleep(1)
ReportPeriodDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/table/tbody/tr/td[2]/img")
ReportPeriodDropDown.click()

time.sleep(1)

ReportYTD = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/div/div/div/ul/li[4]")
ReportYTD.click()

time.sleep(1)

PartGroupsDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/table/tbody/tr/td[2]/img")
PartGroupsDropDown.click()
time.sleep(1)
NoPartGroups = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/div/div/div/ul/li[1]")
NoPartGroups.click()

time.sleep(1)

ShowInDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[2]/img")
ShowInDropDown.click()
time.sleep(1)
ShowInExcel = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/div/div/div/ul/li[3]")
ShowInExcel.click()

iframeSL = browser.find_element(By.ID, "frmGrids")
browser.switch_to.frame(iframeSL)

SL = browser.find_element(By.XPATH, "/html/body/form/div[3]/table/tbody/tr[1]/td[1]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/div/table/tbody/tr[1]/td[1]/table/tbody[2]/tr/td/div[2]/table/tbody/tr/td[4]")
SL.click()
time.sleep(2)

browser.switch_to.default_content()

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

Generate = browser.find_element(By.ID, "btnGenerate__4")
Generate.click()

time.sleep(5)
wait('IkoManagementKPI_Summary.xls','GV.xls')
browser.quit()

############################ GV - EXCEL SECTION ###############################

i = 2
while i/q < 90:
    wb = xw.Book('GV.xls').sheets[0]
    
    i = i + 1
    if wb['B' + str(i)].value == 'OEE (%)':
        LINE = wb['B' + str(i-9)].value
        OEEs = wb.range('C' + str(i) + ':Q' + str(i)).value
        q=q+1
        wb = xw.Book('OEE Summary.xlsx').sheets[0]
        wb['B' + str(q)].value = LINE
        wb['A' + str(q)].value = 'GV'
        wb.range('C' + str(q) + ':Q' + str(q)).value = OEEs
        wb = xw.Book('OEE Summary.xlsx')
        wb.save()


wb = xw.Book('GV.xls')
wb.app.quit()


############################ GI - WEBSITE SECTION ###############################


browser = webdriver.Edge()
browser.get('http://nsikomadiis1/Prodac/Reports/ProdacReportsMain.aspx')
browser.maximize_window()

time.sleep(3)

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

GoTo = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/div/div[2]/ul/li/ul/li[1]/ul/li[3]/a")
GoTo.click()

time.sleep(3)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

FromDate = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/input")
FromDate.click()
FromDate.send_keys('01/01/2023')

GroupingDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/table/tbody/tr/td[2]/img")
GroupingDropDown.click()
time.sleep(1)
GroupingMonth = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/div/div/div/ul/li[2]")
GroupingMonth.click()
time.sleep(1)
ReportPeriodDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/table/tbody/tr/td[2]/img")
ReportPeriodDropDown.click()

time.sleep(1)

ReportYTD = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/div/div/div/ul/li[4]")
ReportYTD.click()

time.sleep(1)

PartGroupsDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/table/tbody/tr/td[2]/img")
PartGroupsDropDown.click()
time.sleep(1)
NoPartGroups = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/div/div/div/ul/li[1]")
NoPartGroups.click()

time.sleep(1)

ShowInDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[2]/img")
ShowInDropDown.click()
time.sleep(1)
ShowInExcel = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/div/div/div/ul/li[3]")
ShowInExcel.click()

iframeSL = browser.find_element(By.ID, "frmGrids")
browser.switch_to.frame(iframeSL)

All = browser.find_element(By.ID, "grdDepartment_chk")
All.click()
time.sleep(2)

browser.switch_to.default_content()

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

Generate = browser.find_element(By.ID, "btnGenerate__4")
Generate.click()

time.sleep(5)
wait('IkoManagementKPI_Summary.xls','GI.xls')
browser.quit()

############################ GI - EXCEL SECTION ###############################


i = 2
while i/q < 90:
    wb = xw.Book('GI.xls').sheets[0]
    
    i = i + 1
    if wb['B' + str(i)].value == 'OEE (%)':
        LINE = wb['B' + str(i-9)].value
        OEEs = wb.range('C' + str(i) + ':Q' + str(i)).value
        q=q+1
        wb = xw.Book('OEE Summary.xlsx').sheets[0]
        wb['B' + str(q)].value = LINE
        wb['A' + str(q)].value = 'GI'
        wb.range('C' + str(q) + ':Q' + str(q)).value = OEEs
        wb = xw.Book('OEE Summary.xlsx')
        wb.save()


wb = xw.Book('GI.xls')
wb.app.quit()



############################ GE - WEBSITE SECTION ###############################


browser = webdriver.Edge()
browser.get('http://igashiis1.na.iko/ProdacNewWeb/Reports/ProdacReportsMain.aspx')
browser.maximize_window()

time.sleep(3)

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

GoTo = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/div/div[2]/ul/li/ul/li[1]/ul/li[3]/a")
GoTo.click()

time.sleep(3)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

FromDate = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/input")
FromDate.click()
FromDate.send_keys('01/01/2023')

GroupingDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/table/tbody/tr/td[2]/img")
GroupingDropDown.click()
time.sleep(1)
GroupingMonth = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/div/div/div/ul/li[2]")
GroupingMonth.click()
time.sleep(1)
ReportPeriodDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/table/tbody/tr/td[2]/img")
ReportPeriodDropDown.click()

time.sleep(1)

ReportYTD = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/div/div/div/ul/li[4]")
ReportYTD.click()

time.sleep(1)

PartGroupsDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/table/tbody/tr/td[2]/img")
PartGroupsDropDown.click()
time.sleep(1)
NoPartGroups = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/div/div/div/ul/li[1]")
NoPartGroups.click()

time.sleep(1)

ShowInDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[2]/img")
ShowInDropDown.click()
time.sleep(1)
ShowInExcel = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/div/div/div/ul/li[3]")
ShowInExcel.click()

iframeSL = browser.find_element(By.ID, "frmGrids")
browser.switch_to.frame(iframeSL)

All = browser.find_element(By.ID, "grdDepartment_chk")
All.click()
time.sleep(2)

browser.switch_to.default_content()

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

Generate = browser.find_element(By.ID, "btnGenerate__4")
Generate.click()

time.sleep(5)
wait('IkoManagementKPI_Summary.xls','GE.xls')
browser.quit()

############################ GE - EXCEL SECTION ###############################



i = 2
while i/q < 90:
    wb = xw.Book('GE.xls').sheets[0]
    
    i = i + 1
    if wb['B' + str(i)].value == 'OEE (%)':
        LINE = wb['B' + str(i-9)].value
        OEEs = wb.range('C' + str(i) + ':Q' + str(i)).value
        q=q+1
        wb = xw.Book('OEE Summary.xlsx').sheets[0]
        wb['B' + str(q)].value = LINE
        wb['A' + str(q)].value = 'GE'
        wb.range('C' + str(q) + ':Q' + str(q)).value = OEEs
        wb = xw.Book('OEE Summary.xlsx')
        wb.save()


wb = xw.Book('GE.xls')
wb.app.quit()

############################ GM - WEBSITE SECTION ###############################


browser = webdriver.Edge()
browser.get('http://nsighiriviis1/Prodac/Reports/ProdacReportsMain.aspx')
browser.maximize_window()

time.sleep(3)

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

GoTo = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/div/div[2]/ul/li/ul/li[1]/ul/li[2]/a")
GoTo.click()
time.sleep(3)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

FromDate = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/input")
FromDate.click()
FromDate.send_keys('01/01/2023')

GroupingDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/table/tbody/tr/td[2]/img")
GroupingDropDown.click()
time.sleep(1)
GroupingMonth = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/div/div/div/ul/li[2]")
GroupingMonth.click()
time.sleep(1)
ReportPeriodDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/table/tbody/tr/td[2]/img")
ReportPeriodDropDown.click()

time.sleep(1)

ReportYTD = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/div/div/div/ul/li[4]")
ReportYTD.click()

time.sleep(1)

PartGroupsDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/table/tbody/tr/td[2]/img")
PartGroupsDropDown.click()
time.sleep(1)
NoPartGroups = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/div/div/div/ul/li[1]")
NoPartGroups.click()

time.sleep(1)

ShowInDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[2]/img")
ShowInDropDown.click()
time.sleep(1)
ShowInExcel = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/div/div/div/ul/li[3]")
ShowInExcel.click()

iframeSL = browser.find_element(By.ID, "frmGrids")
browser.switch_to.frame(iframeSL)

SL = browser.find_element(By.XPATH, "/html/body/form/div[3]/table/tbody/tr[1]/td[1]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/div/table/tbody/tr[1]/td[1]/table/tbody[2]/tr/td/div[2]/table/tbody/tr/td[4]")
SL.click()
time.sleep(2)

browser.switch_to.default_content()

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

Generate = browser.find_element(By.ID, "btnGenerate__4")
Generate.click()

time.sleep(5)
wait('IkoManagementKPI_Summary.xls','GM.xls')
browser.quit()

############################ GM - EXCEL SECTION ###############################



i = 2
while i/q < 90:
    wb = xw.Book('GM.xls').sheets[0]
    
    i = i + 1
    if wb['B' + str(i)].value == 'OEE (%)':
        LINE = wb['B' + str(i-9)].value
        OEEs = wb.range('C' + str(i) + ':Q' + str(i)).value
        q=q+1
        wb = xw.Book('OEE Summary.xlsx').sheets[0]
        wb['B' + str(q)].value = LINE
        wb['A' + str(q)].value = 'GM'
        wb.range('C' + str(q) + ':Q' + str(q)).value = OEEs
        wb = xw.Book('OEE Summary.xlsx')
        wb.save()


wb = xw.Book('GM.xls')
wb.app.quit()


############################ GP - WEBSITE SECTION ###############################


browser = webdriver.Edge()
browser.get('http://nscrcbrmiis1/Prodac/Reports/ProdacReportsMain.aspx')
browser.maximize_window()

time.sleep(3)

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

GoTo = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/div/div[2]/ul/li/ul/li[1]/ul/li[2]/a")
GoTo.click()
time.sleep(3)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

FromDate = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/input")
FromDate.click()
FromDate.send_keys('01/01/2023')

GroupingDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/table/tbody/tr/td[2]/img")
GroupingDropDown.click()
time.sleep(1)
GroupingMonth = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/div/div/div/ul/li[2]")
GroupingMonth.click()
time.sleep(1)
ReportPeriodDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/table/tbody/tr/td[2]/img")
ReportPeriodDropDown.click()

time.sleep(1)

ReportYTD = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/div/div/div/ul/li[4]")
ReportYTD.click()

time.sleep(1)

PartGroupsDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/table/tbody/tr/td[2]/img")
PartGroupsDropDown.click()
time.sleep(1)
NoPartGroups = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/div/div/div/ul/li[1]")
NoPartGroups.click()

time.sleep(1)

ShowInDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[2]/img")
ShowInDropDown.click()
time.sleep(1)
ShowInExcel = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/div/div/div/ul/li[3]")
ShowInExcel.click()

iframeSL = browser.find_element(By.ID, "frmGrids")
browser.switch_to.frame(iframeSL)

SL = browser.find_element(By.XPATH, "/html/body/form/div[3]/table/tbody/tr[1]/td[1]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/div/table/tbody/tr[1]/td[1]/table/tbody[2]/tr/td/div[2]/table/tbody/tr/td[4]")
SL.click()
time.sleep(2)

browser.switch_to.default_content()

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

Generate = browser.find_element(By.ID, "btnGenerate__4")
Generate.click()

time.sleep(5)
wait('IkoManagementKPI_Summary.xls','GP.xls')
browser.quit()

############################ GP - EXCEL SECTION ###############################


i = 2
while i/q < 90:
    wb = xw.Book('GP.xls').sheets[0]
    
    i = i + 1
    if wb['B' + str(i)].value == 'OEE (%)':
        LINE = wb['B' + str(i-9)].value
        OEEs = wb.range('C' + str(i) + ':Q' + str(i)).value
        q=q+1
        wb = xw.Book('OEE Summary.xlsx').sheets[0]
        wb['B' + str(q)].value = LINE
        wb['A' + str(q)].value = 'GP'
        wb.range('C' + str(q) + ':Q' + str(q)).value = OEEs
        wb = xw.Book('OEE Summary.xlsx')
        wb.save()


wb = xw.Book('GP.xls')
wb.app.quit()


############################ BL - WEBSITE SECTION ###############################


browser = webdriver.Edge()
browser.get('http://nsblmhgriis1/PRODAC/Reports/ProdacReportsMain.aspx')
browser.maximize_window()

time.sleep(3)

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

GoTo = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/div/div[2]/ul/li/ul/li[1]/ul/li[2]/a")
GoTo.click()
time.sleep(3)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

FromDate = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/input")
FromDate.click()
FromDate.send_keys('01/01/2023')

GroupingDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/table/tbody/tr/td[2]/img")
GroupingDropDown.click()
time.sleep(1)
GroupingMonth = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/div/div/div/ul/li[2]")
GroupingMonth.click()
time.sleep(1)
ReportPeriodDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/table/tbody/tr/td[2]/img")
ReportPeriodDropDown.click()

time.sleep(1)

ReportYTD = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/div/div/div/ul/li[4]")
ReportYTD.click()

time.sleep(1)

PartGroupsDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/table/tbody/tr/td[2]/img")
PartGroupsDropDown.click()
time.sleep(1)
NoPartGroups = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/div/div/div/ul/li[1]")
NoPartGroups.click()

time.sleep(1)

ShowInDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[2]/img")
ShowInDropDown.click()
time.sleep(1)
ShowInExcel = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/div/div/div/ul/li[3]")
ShowInExcel.click()

iframeSL = browser.find_element(By.ID, "frmGrids")
browser.switch_to.frame(iframeSL)

SL = browser.find_element(By.XPATH, "/html/body/form/div[3]/table/tbody/tr[1]/td[1]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/div/table/tbody/tr[1]/td[1]/table/tbody[2]/tr/td/div[2]/table/tbody/tr/td[4]")
SL.click()
time.sleep(2)

browser.switch_to.default_content()

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

Generate = browser.find_element(By.ID, "btnGenerate__4")
Generate.click()

time.sleep(5)
wait('IkoManagementKPI_RL_ISOBOARD_Summary.xls','BL.xls')
browser.quit()

############################ BL - EXCEL SECTION ###############################

i = 2
while i/q < 90:
    wb = xw.Book('BL.xls').sheets[0]
    
    i = i + 1
    if wb['B' + str(i)].value == 'OEE (%)':
        LINE = wb['B' + str(i-9)].value
        OEEs = wb.range('C' + str(i) + ':Q' + str(i)).value
        q=q+1
        wb = xw.Book('OEE Summary.xlsx').sheets[0]
        wb['B' + str(q)].value = LINE
        wb['A' + str(q)].value = 'BL'
        wb.range('C' + str(q) + ':Q' + str(q)).value = OEEs
        wb = xw.Book('OEE Summary.xlsx')
        wb.save()


wb = xw.Book('BL.xls')
wb.app.quit()

############################ BLTPO - WEBSITE SECTION ###############################


browser = webdriver.Edge()
browser.get('http://nsblmhgriis1/PRODAC/Reports/ProdacReportsMain.aspx')
browser.maximize_window()

time.sleep(3)

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

GoTo = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/div/div[2]/ul/li/ul/li[1]/ul/li[5]/a")
GoTo.click()
time.sleep(3)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

FromDate = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/input")
FromDate.click()
FromDate.send_keys('01/01/2023')

GroupingDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/table/tbody/tr/td[2]/img")
GroupingDropDown.click()
time.sleep(1)
GroupingMonth = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div/div/div/div/div/ul/li[2]")
GroupingMonth.click()
time.sleep(1)
ReportPeriodDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/table/tbody/tr/td[2]/img")
ReportPeriodDropDown.click()

time.sleep(1)

ReportYTD = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[3]/div/div/div/div/div/ul/li[4]")
ReportYTD.click()

time.sleep(1)

PartGroupsDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/table/tbody/tr/td[2]/img")
PartGroupsDropDown.click()
time.sleep(1)
NoPartGroups = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[4]/div/div/div/div/div/ul/li[1]")
NoPartGroups.click()

time.sleep(1)

ShowInDropDown = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[2]/img")
ShowInDropDown.click()
time.sleep(1)
ShowInExcel = browser.find_element(By.XPATH, "/html/body/form/table[1]/tbody/tr[3]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/div/table/tbody/tr[1]/td/div/div/div/div/div/ul/li[3]")
ShowInExcel.click()

iframeSL = browser.find_element(By.ID, "frmGrids")
browser.switch_to.frame(iframeSL)

SL = browser.find_element(By.XPATH, "/html/body/form/div[3]/table/tbody/tr[1]/td[1]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/div/table/tbody/tr[1]/td[1]/table/tbody[2]/tr/td/div[2]/table/tbody/tr/td[4]")
SL.click()
time.sleep(2)

browser.switch_to.default_content()

iframeinitial = browser.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/div/div/div/div[1]/iframe")
browser.switch_to.frame(iframeinitial)

iframeReports = browser.find_element(By.ID, "frmReports")
browser.switch_to.frame(iframeReports)

Generate = browser.find_element(By.ID, "btnGenerate__4")
Generate.click()

time.sleep(5)
wait('IkoManagementKPI_RL_NA_Summary.xls','BLTPO.xls')
browser.quit()

############################ BL - EXCEL SECTION ###############################

i = 2
while i/q < 90:
    wb = xw.Book('BLTPO.xls').sheets[0]
    
    i = i + 1
    if wb['B' + str(i)].value == 'OEE (%)':
        LINE = wb['B' + str(i-9)].value
        OEEs = wb.range('C' + str(i) + ':Q' + str(i)).value
        q=q+1
        wb = xw.Book('OEE Summary.xlsx').sheets[0]
        wb['B' + str(q)].value = LINE
        wb['A' + str(q)].value = 'BLTPO'
        wb.range('C' + str(q) + ':Q' + str(q)).value = OEEs
        wb = xw.Book('OEE Summary.xlsx')
        wb.save()


wb = xw.Book('BLTPO.xls')
wb.app.quit()

print("See OEE Summary Excel Sheet")


