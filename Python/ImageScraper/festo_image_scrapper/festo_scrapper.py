# -*- coding: utf-8 -*-
"""
Created on July 20, 2023

v0.1.0

@author: Desumai

- Functions for image scrapping from Festo.
"""

from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver
import scrapper





def find_search_box_festo(driver: WebDriver):
    """
    Finds the search bar/box element on Festo's webpage
    @params
        driver - Required: the webdriver objet that is on the page (selenium.webdriver.remote.webelement.WebDriver)
    """
    return driver.find_element(
        By.XPATH, '//*[@id="first-search-input"]'
    )



def get_product_img_via_searchbox_festo(
     driver: WebDriver, searchTerm: str
):
    """
    Gets image urls by searching with a search bar webelement

    @params
        driver - the (Chrome) WebDriver object (selenium.webdriver.remote.webdriver.WebDriver)
        searchTerm - the text to search (str)
    """
    driver.implicitly_wait(1)
    search = find_search_box_festo(driver=driver)
    scrapper.new_search(search, searchTerm)
    driver.implicitly_wait(4)
    try:
        return get_product_image_URL_festo(driver)
    except Exception as e:
        # print(e)
        return None

    
def get_product_image_URL_festo(driver: WebDriver):
    """
    Gets the image url of the first product found after a search on festo.com/us

    @params
        driver - (Chrome) WebDriver object (selenium.webdriver.remote.webdriver.WebDriver)
    """
    product = driver.find_element(By.CLASS_NAME, "single-product-container--tD974")
    product_image = product.find_element(By.XPATH, "//div/div[1]/img")
    return product_image.get_attribute("src")

def handle_cookies_popup_festo(driver:WebDriver):
    '''
    Handles cookies popup on festo.com if it appears.

    @params
        driver - (Chrome) WebDriver object (selenium.webdriver.remote.webdriver.WebDriver)
    '''
    try:
        #find popup and 'accept all cookies' button
        driver.implicitly_wait(2)
        popup = driver.find_element(By.CLASS_NAME, "overlay--yeLyO")
        popupBtn = popup.find_element(By.XPATH, "//div/div/menu/button[2]")
        #click 'accept all cookies' button
        popupBtn.click()
    except Exception as e:
        #no cookies popup
        #print('no cookies popup')
        pass

def load_festo_CSV_file_to_dict(csvPATH: str):
    """
    Imports an input CSV file and processes it to a dict.
    Format for entries is {maxmo #: [list of search terms]}
    Returns the dict. Note that the csv file must comma delimited.

    @params
        csvPATH - the file path to the input csv file (str)
    """
    import csv

    itemDict = {}
    with open(csvPATH) as csvFile:
        reader = list(csv.reader(csvFile, delimiter=",", quotechar='"'))
        reader.pop(0)  # remove header row
        try:
            for row in range(0, len(reader)):
                try:
                    # get maximo num
                    maximoNum = reader[row][0]
                    if len(maximoNum) != 7:
                        raise Exception("Maximo number is not 7 digits long")
                    maximoNum = int(maximoNum)
                    if maximoNum < 9000000:
                        raise Exception("Maximo number must start with a 9")
                    # get festo search terms
                    searchTerms = []
                    for searchTerm in reader[row][1:]:
                        if(searchTerm != ''):
                            searchTerms.append(searchTerm)
                    # put entry in dict
                    if maximoNum in itemDict.keys():
                        raise Exception(
                            "Maximo number in row " + row + " is a duplicate"
                        )
                    itemDict[maximoNum] = searchTerms
                except Exception as e:
                    print(
                        "Error occured in line "
                        + row
                        + " when loading csv file:\n\t"
                        + e
                    )
        except Exception as e:
            print(e)
    return itemDict