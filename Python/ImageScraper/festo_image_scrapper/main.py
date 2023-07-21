# -*- coding: utf-8 -*-
"""
Created on July 20, 2023

v0.1.0

@author: Desumai

- Selenium image scrapping script for Festo. Uses text input search to find items instead of URLs.
"""

import os
import festo_scrapper as fs
import scrapper

#initialize variables
URL = "https://www.festo.com/us/en/search/?tab=PRODUCTS&q="
savePATH = None # the path to the directory to save the images to
imageURLDict = {}  # dict of maximo #s to found image urls
errorDict = (
    {}
)  # dict of the items that the program failed to find URLs for. Key = maximo #, value = festo description
itemDict = None  # dict of the items to search. Key = maximo #, value = festo description

#get user input csv file
print("Choose a csv input file to import (comma delimited)")
csvPATH = None
try:
    csvPATH = scrapper.get_CSV_path()
except ValueError as e:
    print(e)
    print("Could not load file")
if csvPATH is None:
    scrapper.close_program()

#process csv file
try:
    itemDict = fs.load_festo_CSV_file_to_dict(csvPATH)
except FileNotFoundError as e:
    print("Could not find file '" + e.filename + "'")
    scrapper.close_program()
except Exception as e:
    print(e)

print("csv file loaded.")

#choose save directory
print("Choose directory to save images to")
try:
    savePATH = scrapper.get_save_path(defaultSavePATH=os.getcwd)
except Exception as e:
    print(e)
    print("Could not save to this location")
if savePATH is None:
    scrapper.close_program()

#initialize scraper

print("Starting up Selenium...")
driver = None
try:
    driver = scrapper.initialize_driver(URL)
    fs.handle_cookies_popup_festo(driver)
except Exception as e:
    print(e)
    print(
        "Could not initialize Selenium Chromedriver.\nMake sure you have the proper chrome driver installed onto PATH"
    )
    scrapper.close_program()

#search for image URLs
print("Finding image URLs...")
# get image URLs
numOfItems = len(itemDict.items())
scrapper.printProgressBar(
    0, numOfItems, prefix="Progress:", suffix="Complete", length=50
)
for i, entry in enumerate(itemDict.items()):
    tempURL = None
    searchTerms = entry[1]
    for searchTerm in searchTerms:
        try:
            tempURL = fs.get_product_img_via_searchbox_festo(driver, searchTerm
            )
            if (tempURL is not None):
                break
        except Exception as e:
            # print(e)
            pass
    if tempURL is not None:
        imageURLDict[entry[0]] = tempURL
    else:
        errorDict[entry[0]] = entry[1]
    scrapper.printProgressBar(
        i + 1,
        numOfItems,
        prefix="Progress:",
        suffix="Complete " + scrapper.get_spinner_char(i),
        length=50,
    )

print("Completed.\n\nDownloading images...")

# downloading images
numOfItems = len(imageURLDict.items())
if numOfItems < 1:
    raise Exception("No images to download")
    # TODO: handle no images to download
scrapper.printProgressBar(
    0, numOfItems, prefix="Progress:", suffix="Complete", length=50
)
for i, entry in enumerate(imageURLDict.items()):
    success = scrapper.save_url_image_to_file(entry[1], str(entry[0]), savePATH)
    if not success:
        errorDict[entry[0]] = entry[1]
    scrapper.printProgressBar(
        i + 1,
        numOfItems,
        prefix="Progress:",
        suffix="Complete " + scrapper.get_spinner_char(i),
        length=50,
    )
print("Completed.\nDownloaded to '" + savePATH + "'\n\n")

# print results
print("Image scrapping complete.")
print(
    str(len(itemDict.items()) - len(errorDict.items()))
    + " succesful, "
    + str(len(errorDict.items()))
    + " failed."
)
print(str(len(itemDict.items())) + " items total.")

# save error log if needed
if len(errorDict.items()) > 0:
    scrapper.save_error_dict(savePATH, errorDict)
    print("List of failed items saved to '" + savePATH + "failed_items.txt'.")

print()
scrapper.close_driver(driver)
