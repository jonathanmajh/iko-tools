# -*- coding: utf-8 -*-
"""
Created on July 12, 2023

@author: Desumai

v0.1.0
- Functions for image scrapping from SKF.
- Functions can be reformatted and reused for image scrapping other sites
- Not all functions are used
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.remote.webdriver import WebDriver
import requests, easygui, urllib.parse
from selenium.webdriver import ChromeOptions


def initialize_driver(URL: str):
    """
    Creates and returns a driver using from the chromedriver on OS PATH.
    Initializes a headless browser

    @param
        URL - the URL of the website to search (str)
    """
    # hide browser window (headless browser)
    options = ChromeOptions()
    options.headless = True
    options.add_argument("log-level=3")  # hide non-fatal error messages from selenium
    # intialize chromedriver
    driver = webdriver.Chrome(options=options)
    driver.get(URL)
    driver.implicitly_wait(1)
    return driver


def split_SKF_URL(URL: str, searchComponent: str):
    """
    Splits the SKF URL into a list of 2 strings, seperated by the end of the search component. Used to create search URLs
    @param
        URL - the SKL URL (str)
        searchComponent - the search term to split the SKL URL around (str)

    example:
        split_SKF_URL(URL="https://website.com/search?q=&lan=en&include=all", searchComponent="search?q=") => ["https://website.com/search?q=","&lan=en&include=all"]
    """
    searchLocation = URL.index(searchComponent)
    if searchLocation == -1:
        raise Exception("Improper URL, does not contain '" + searchComponent + "'.")
    searchLocation += len(searchComponent)
    return [URL[:searchLocation], URL[searchLocation:]]


def create_search_URL(URLComponents: list, searchTerm: str = ""):
    """
    Creates and returns a URL that will search a given search term on a website domain

    @params
        URLComponents - list<str> of length 2 that contains the first half of the website in index [0]
                        and the second half in index [1] (list<str>)
        searchTerm - the term to search (str)
    """
    return URLComponents[0] + urllib.parse.quote(searchTerm) + URLComponents[1]


def get_img_URL_via_URL_search_SKF(
    URLComponents: list, searchTerm: str, driver: WebDriver
):
    """
    Gets the image URL by searching using URLs on driver.get

    @params
        URLComponents - the two halves of the website domain URL (list<str>)
        searchTerm - the search term to find the item to scrap. Should be a manufacturer
                    number. (str)
        driver - the (Chrome) WebDriver object (selenium.webdriver.remote.webdriver.WebDriver)
    """
    searchURL = create_search_URL(URLComponents=URLComponents, searchTerm=searchTerm)
    driver.get(searchURL)
    driver.implicitly_wait(6)
    try:
        return get_product_image_URL_SKF(driver)
    except Exception as e:
        # print(e)
        return None


def find_search_box_SKF(driver: WebDriver):
    """
    Finds the search bar/box element on SFK's webpage
    @params
        driver - Required: the webdriver objet that is on the page (selenium.webdriver.remote.webelement.WebDriver)
    """
    return driver.find_element(
        By.XPATH, '//*[@id="cid-480811"]/div[2]/div/search-input/form/input'
    )


def clear_field(webElement: WebElement):
    """
    Clears the input field of a webElement
    @params
        webElement -  the WebElement whose input field to clear (selenium.webdriver.remote.webelement.WebElement)
    """
    webElement.send_keys(Keys.CONTROL + "a")
    webElement.send_keys(Keys.DELETE)


def new_search(webElement: WebElement, searchTerm: str):
    """
    Does a new search on the webElement. Remember to wait for webpage to load

    @params
        webElement - the input-accepting WebElement to search on (selenium.webdriver.remote.webelement.WebElement)
        searchTerm - what to search
    """
    clear_field(webElement)
    webElement.send_keys(searchTerm)
    webElement.send_keys(Keys.RETURN)


def get_product_image_URL_SKF(driver: WebDriver):
    """
    Gets the image url of the first product found after a search on SKF.com

    @params
        driver - (Chrome) WebDriver object (selenium.webdriver.remote.webdriver.WebDriver)
    """
    product = driver.find_element(By.CLASS_NAME, "product-image")
    product_image = product.find_element(By.CLASS_NAME, "card-img-top")
    return product_image.get_attribute("src")


def get_product_img_via_searchbox_SKF(
    searchBar: WebElement, driver: WebDriver, searchTerm: str
):
    """
    Gets image urls by searching with a search bar webelement

    @params
        searchBar - the search bar WebElement to search on (selenium.webdriver.remote.webelement.WebElement)
        driver - the (Chrome) WebDriver object (selenium.webdriver.remote.webdriver.WebDriver)
        searchTerm - the text to search (str)
    """
    search = find_search_box_SKF(driver=driver)
    new_search(search, searchTerm)
    driver.implicitly_wait(4)
    try:
        return get_product_image_URL_SKF(driver)
    except Exception as e:
        # print(e)
        return None


def close_program():
    """Close the program when user presses ENTER"""
    input("Program terminated. Press ENTER to close...")
    exit()


def close_driver(driver: WebDriver):
    """Close the WebDriver and program"""
    driver.quit()
    close_program()


def get_image_file_extension_from_URL(imgURL: str):
    """
    Gets the file extension for a image URL
    @params
        imgURL - the image URL (str)
    """
    extIndex = imgURL.rindex(".")
    if extIndex < 1 or extIndex <= imgURL.rindex(
        "/"
    ):  # must be last period after last backslash
        raise Exception("Image URL '" + imgURL + "' is not valid")
    return imgURL[extIndex:]


def save_url_image_to_file(
    imgURL: str, fileName: str, filePATH: str, fileExt: str = None
):
    """
    Downloads an image from an URL. Returns true if sucessful download, else returns false
    @params
        imgURL - the URL of the image to download (str)
        fileName - the file name to save the image as. Should be the maximo # (str)
        filePATH - the file path to save to (str)
        fileExt - the file extension to use. If None or left empty, uses the file extension from the URL (str)
    """
    try:
        img = requests.get(imgURL)
        if fileExt is None:
            fileExt = get_image_file_extension_from_URL(imgURL=imgURL)
        if img.status_code == 200:
            with open(filePATH + fileName + fileExt, "wb") as handler:
                handler.write(img.content)
            return True
        else:
            return False
    except Exception as e:
        # print(e)
        return False


def get_CSV_path():
    """
    Opens a file picker dialog (easygui) for the user to select an input CSV file.
    Returns None if the user cancels the dialog.
    """
    importPATH = easygui.fileopenbox(
        msg="Please Select CSV File",
        title="Open File",
        default="c:\data\det\*.csv",
        filetypes="*.csv",
        multiple=False,
    )
    # ensure file is csv
    if type(importPATH) == type(""):
        extIndex = importPATH.rindex(".")
        extStr = importPATH[extIndex + 1 :]
        if (extIndex <= 0) or (extStr.lower() != "csv"):
            raise ValueError("File '" + importPATH + "' is not of type .csv")

    return importPATH


def get_save_path(defaultSavePATH: str = None):
    """
    Opens a file picker dialog (easygui) for the user to select a directory to save
    the output images to. Returns None if the user cancels the dialog.
    @params
        defaultSavePATH - the file path to the default directory to save to. Optional. (str)

    """
    savePATH = easygui.diropenbox(title="Save To", default=defaultSavePATH)
    if savePATH is not None:
        savePATH = savePATH + "\\"
    return savePATH


def load_SFK_CSV_file_to_dict(csvPATH: str):
    """
    Imports an input CSV file and processes it to a dict.
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
                    # get skf code
                    SFKNum = reader[row][1]
                    # put entry in dict
                    if maximoNum in itemDict.keys():
                        raise Exception(
                            "Maximo number in row " + row + " is a duplicate"
                        )
                    itemDict[maximoNum] = SFKNum
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


# Print iterations progress. from https://stackoverflow.com/questions/3173320/text-progress-bar-in-terminal-with-block-characters
def printProgressBar(
    iteration,
    total,
    prefix="",
    suffix="",
    decimals=1,
    length=100,
    fill="█",
    printEnd="\r",
):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + "-" * (length - filledLength)
    print(f"\r{prefix} |{bar}| {percent}% {suffix}", end=printEnd)
    # Print New Line on Complete
    if iteration == total:
        print()


def save_error_dict(savePATH: str, errorDict: list):
    """
    Saves a list of items that the program failed to download images
    for into 'failed_items.txt' in the selected save directory.

    @params
        savePATH - path of the save directory (str)
        errorDict - jagged list of the failed items. Each entry is a list<str> of length 2 where
                    [0] is the maximo # and [1] is the item info ()
    """
    with open(savePATH + "failed_items.txt", "w") as file:
        for entry in errorDict.items():
            file.write(str(entry[0]) + "," + str(entry[1]))
            file.write("\n")


def get_spinner_char(i: int, passCount: int = 1, spinnerChars: str = "|/-\\"):
    """
    Gets the character for the current iteration of the spinner
    @params
        i - current iteration (int)
        passCount - how many iterations to pass before changing spinner character. Default is 1 (int)
        spinnerChars - string of the characters to use for the spinner. Default is |/-\ (str)
    """
    pos = round(i / passCount) % len(spinnerChars)
    return spinnerChars[pos]
