# -*- coding: utf-8 -*-
"""
Created on July 12, 2023

@author: Desumai

v0.1.0
- Image scrapper for SKF. Scraps the images from SKF's site according to input comma delimited csv file.
- CSV file should be 2 columns by x rows with a header. Entries are written as: [maximo #],[SKF code]
- Saves images into a folder, each image is name according to its maximo number
- Uses URLs to search for items instead of search bar
- See ..\\dist for release. Remember to update dist with pyinstaller when modifying source code
"""

import skfscapper
import os

# initialize variables
URL = "https://www.skf.com/group/search-results?q=&searcher=products&site=307&language=en&tridion_target=live&tridion_version=3&language_preset=English"
searchComponent = "?q="  # the search query part of the URL
savePATH = None  # the path to the directory to save the images in
URLComponents = skfscapper.split_SKF_URL(
    URL=URL, searchComponent=searchComponent
)  # list of url components, to make searching for items easier
imageURLDict = {}  # dict of maximo #s to found image urls
errorDict = (
    {}
)  # dict of the items that the program failed to find URLs for. Key = maximo #, value = SKF description
itemDict = None  # dict of the items to search. Key = maximo #, value = SKF description

# get input csv file
print("Choose a csv input file to import (comma delimited)")
csvPATH = None
try:
    csvPATH = skfscapper.get_CSV_path()
except ValueError as e:
    print(e)
    print("Could not load file")
if csvPATH is None:
    skfscapper.close_program()

# import csv file:
try:
    itemDict = skfscapper.load_SFK_CSV_file_to_dict(csvPATH)
except FileNotFoundError as e:
    print("Could not find file '" + e.filename + "'")
    skfscapper.close_program()
except Exception as e:
    print(e)

print("csv file loaded.")

# set save directory
print("Choose directory to save images to")
try:
    savePATH = skfscapper.get_save_path(defaultSavePATH=os.getcwd)
except Exception as e:
    print(e)
    print("Could not save to this location")
if savePATH is None:
    skfscapper.close_program()


print("Starting up Selenium...")
driver = None
try:
    driver = skfscapper.initialize_driver(URL)
except Exception as e:
    print(e)
    print(
        "Could not initialize Selenium Chromedriver.\nMake sure you have the proper chrome driver installed onto PATH"
    )
    skfscapper.close_program()


print("Finding image URLs...")
# get image URLs
numOfItems = len(itemDict.items())
skfscapper.printProgressBar(
    0, numOfItems, prefix="Progress:", suffix="Complete", length=50
)
for i, entry in enumerate(itemDict.items()):
    tempURL = None
    try:
        tempURL = skfscapper.get_img_URL_via_URL_search_SKF(
            URLComponents=URLComponents, driver=driver, searchTerm=entry[1]
        )
    except Exception as e:
        # print(e)
        pass
    if tempURL is not None:
        imageURLDict[entry[0]] = tempURL
    else:
        errorDict[entry[0]] = entry[1]
    skfscapper.printProgressBar(
        i + 1,
        numOfItems,
        prefix="Progress:",
        suffix="Complete " + skfscapper.get_spinner_char(i),
        length=50,
    )

print("Completed.\n\nDownloading images...")


# downloading images
numOfItems = len(imageURLDict.items())
if numOfItems < 1:
    raise Exception("No images to download")
    # TODO: handle no images to download
skfscapper.printProgressBar(
    0, numOfItems, prefix="Progress:", suffix="Complete", length=50
)
for i, entry in enumerate(imageURLDict.items()):
    success = skfscapper.save_url_image_to_file(entry[1], str(entry[0]), savePATH)
    if not success:
        errorDict[entry[0]] = entry[1]
    skfscapper.printProgressBar(
        i + 1,
        numOfItems,
        prefix="Progress:",
        suffix="Complete " + skfscapper.get_spinner_char(i),
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
    skfscapper.save_error_dict(savePATH, errorDict)
    print("List of failed items saved to '" + savePATH + "failed_items.txt'.")

print()
skfscapper.close_driver(driver)
