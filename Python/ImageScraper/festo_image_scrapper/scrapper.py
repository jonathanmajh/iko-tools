# -*- coding: utf-8 -*-
"""
Created on July 20, 2023

v0.1.0

@author: Desumai

- Reusable functions for Selenium (Chrome) image scrappers .
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

def close_program():
    """Close the program when user presses ENTER"""
    input("Program terminated. Press ENTER to close...")
    exit()


def close_driver(driver: WebDriver):
    """Close the WebDriver and program"""
    driver.quit()
    close_program()

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
    imgURL: str, fileName: str, filePATH: str, fileExt: str = None, headers: set = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.82'}
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
        
        img = requests.get(imgURL,headers=headers)
        if fileExt is None:
            fileExt = get_image_file_extension_from_URL(imgURL=imgURL)
        if img.status_code == 200:
            with open(filePATH + fileName + fileExt, "wb") as handler:
                handler.write(img.content)
            return True
        else:
            print(str(img.status_code) + "error")
            return False
    except Exception as e:
        print('Error: ' + str(e))
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
