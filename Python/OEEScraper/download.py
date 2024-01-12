# Script for compiling OEE is in ReliabilityAwardTrackingYYYY-MM.xlsm file

import time, keyboard

from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

sites = [
['IKO Brampton','http://nsikobrmiis1n/PRODAC/Reports/ProdacReportsMain.aspx', 'Management KPI'],
['IG Brampton','http://nsigbrmiis1/PRODAC/Reports/ProdacReportsMain.aspx', 'RL_Management KPI'],
['Kankakee','http://ikokankiis1/ProdacNewWeb/Reports/ProdacReportsMain.aspx', 'SL_Management KPI'],
['Hagerstown','http://nsblmhgriis1/PRODAC/Reports/ProdacReportsMain.aspx', 'ISO_Management KPI'],
['Hawkesbury','http://ikohawkiis2/ProdacNewWeb/Reports/ProdacReportsMain.aspx', 'Management KPI'],
['Calgary','http://ikocalgiis1/ProdacNewWeb/Reports/ProdacReportsMain.aspx', 'Management KPI'],
['Sumas','http://ikosumiis1/ProdacNewWeb/Reports/ProdacReportsMain.aspx', 'SL_Management KPI'],
['Sylacauga','http://nsikosyliis1/Prodac/Reports/ProdacReportsMain.aspx', 'Management KPI'],
['Hillsboro','http://nsikohiliis1.na.iko/PRODAC/Reports/ProdacReportsMain.aspx', 'Management KPI'],
['Madoc','http://nsikomadiis1/Prodac/Reports/ProdacReportsMain.aspx', 'Management KPI'], #requires extra group by Department
['Ashcroft','http://igashiis1.na.iko/ProdacNewWeb/Reports/ProdacReports.aspx', 'Management KPI'],  #requires extra group by Department
['IG High River','http://nsighiriviis1/Prodac/Reports/ProdacReportsMain.aspx', 'Management KPI'],
['CRC Brampton','http://nscrcbrmiis1/Prodac/Reports/ProdacReportsMain.aspx', 'Management KPI'],
# report include all lines, not sure why it's seperated out
# ['IG Brampton','http://nsigbrmiis1/PRODAC/Reports/ProdacReportsMain.aspx', 'B6/Laminator Management KPI'],
# ['Kankakee','http://ikokankiis1/ProdacNewWeb/Reports/ProdacReportsMain.aspx', 'RL_Management KPI'],
# ['Hagerstown','http://nsblmhgriis1/PRODAC/Reports/ProdacReportsMain.aspx', 'TPO_Management KPI'],
# ['Sumas','http://ikosumiis1/ProdacNewWeb/Reports/ProdacReportsMain.aspx', 'RL_Management KPI'],
]


def openAndDownload(link, name):
    driver.get(link)
    iframe = WebDriverWait(driver, timeout=5).until(
        EC.presence_of_element_located((By.XPATH, "//iframe[@title='ReportsMain.aspx']")))
    driver.switch_to.frame(iframe)
    selection = WebDriverWait(driver, timeout=5).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'igdt_ProdacStyle4Node')))

    for element in selection:
        if (name == element.text):
            print(element.text)
            element.click()
            time.sleep(1)
            # driver.switch_to.default_content()
            iframe2 = WebDriverWait(driver, timeout=5).until(EC.presence_of_element_located((By.ID, "frmReports")))
            driver.switch_to.frame(iframe2)
            # select time period
            select = WebDriverWait(driver, timeout=5).until(EC.presence_of_element_located((By.ID, "drpPeriod")))
            select2 = WebDriverWait(select, timeout=5).until(EC.presence_of_element_located((By.TAG_NAME, "input")))
            select2.clear()
            select2.send_keys('Current Year')
            time.sleep(1)
            # select grouping
            select = WebDriverWait(driver, timeout=5).until(EC.presence_of_element_located((By.ID, "ddRV2")))
            select2 = WebDriverWait(select, timeout=5).until(EC.presence_of_element_located((By.TAG_NAME, "input")))
            select2.clear()
            select2.send_keys('Month')
            time.sleep(1)
            # select group by
            select = WebDriverWait(driver, timeout=5).until(EC.presence_of_element_located((By.ID, "ddRV4")))
            select2 = WebDriverWait(select, timeout=5).until(EC.presence_of_element_located((By.TAG_NAME, "input")))
            select2.clear()
            select2.send_keys('No Part Groups')
            time.sleep(1)
            # select show in
            select = WebDriverWait(driver, timeout=5).until(EC.presence_of_element_located((By.ID, "drpViewer")))
            select2 = WebDriverWait(select, timeout=5).until(EC.presence_of_element_located((By.CLASS_NAME, "igdd_ProdacStyle4DropDownButton")))
            select2.click()
            list = WebDriverWait(select, timeout=5).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'igdd_ProdacStyle4ListItem')))
            time.sleep(1)
            for item in list:
                if (item.text == 'Excel'):
                    time.sleep(1)
                    item.click()
            time.sleep(1)
            # submit report btnGenerate__4
            select = WebDriverWait(driver, timeout=5).until(EC.presence_of_element_located((By.ID, "btnGenerate__4")))
            select.click()

driver = webdriver.Chrome()
driver.maximize_window()

for site in sites:
    try:
        print(f'Processing: ${site[0]}')
        openAndDownload(site[1], site[2])
    except:
        pass
    finally:
        driver.switch_to.new_window('tab')

print('New Tabs Complete')
a = keyboard.read_key()