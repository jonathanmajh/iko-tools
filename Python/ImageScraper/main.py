

# -*- coding: utf-8 -*-
"""
Created on Fri Oct 28 10:41:24 2022

@author: wongdere
"""
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from googlesearch import search
import pandas as pd
import urllib
k=0
df2=pd.read_excel('/Users/wongdere/Documents/test.xlsx')
sheet_input=df2['wikipog']
df = pd.DataFrame(columns=['Part number'])
driver = webdriver.Chrome('/Users/wongdere/Downloads/chromedriver')
for inputs in sheet_input:
    #query=inputs
    #for j in search(str(query),tld="co.in",num=10,stop=1,pause=2):
        #searchresult=j
    #strsearcj=str(searchresult)
    #print(strsearcj)
    
    
    url = 'https://www.skf.com/ca/en/search-results?'+str(inputs)
    print(url)
    driver.get(url)
    img=driver.find_element(By.XPATH,'/html/body/div[3]/div[3]/div[5]/div[1]/table[1]/tbody/tr[3]/td/a/img') #img search
    src=img.get_attribute('src')
    urllib.request.urlretrieve(src,str(k)+"test.png")
    #item = driver.find_element(By.CSS_SELECTOR, '[name="q"]').send_keys("webElement")
    #k=0
    #p=0
    #listmultiplier=26
    #inputlist=[0]*listmultiplier
    #for i in range(ord('A'), ord('Z')+1):
        #inputlist[k]=chr(i)
        #print(inputlist[k])
        #k=k+1
    #for u in inputlist:
        #try:
            #input_el = driver.find_element(By.NAME,str(inputlist[p]))
            #td_p_input = input_el.find_element(By.XPATH,'./..').text
        #except NoSuchElementException:
            #print("parent element",inputlist[p],"cannot be found")
        #p=p+1
    input_el = driver.find_element(By.ID,'bodyContent') #text
    #input_el=driver.find_element(By.CLASS_NAME,'swiper-slide')
    td_p_input = input_el.find_element(By.XPATH,'..').text
    #td_p_input = driver.find_element(By.XPATH,"/html/body/div[3]/div/div/div[3]/tridion-product-page-info/main/div/div/div/div/div[1]/div[1]/swiper/div/div[1]").text
    #item= driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[1]/main/div/div/div/div[2]/div[2]/h1').text
    #print(td_p_input)
    item_list = [1]
    #for p in range(len(item)):
        #item_list.append(item[p])
    item_list[0]=td_p_input
    #print(item_list)
    #data_tuples = list(zip(item_list[0]))
    #data_tuples = list(item_list[0])
    data_tuples=item_list
    #print(data_tuples)
    temp_df = pd.DataFrame(data_tuples, columns=['Item'])
    df = df.append(temp_df)
    k=k+1
df.to_excel('/Users/wongdere/Documents/itempictemplate.xlsx')

driver.close()

    
#print(query)

