import requests
import json

LATMAX = 43.90135
LONGMAX = -79.05332
LATMIN = 43.51420
LONGMIN = -79.69945
PAGE = 2
# each response only gives 12 results change the page variable for more responses

# 1. use chrome's user-agent to make a request to the homepage to get session cookies
# this step should only need to be done once per session
header = {
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36',
}
req = requests.get('https://www.realtor.ca/', headers=header)

# 2. use obtained cookies to make a post request to the Properties Listings API
# repeat step 2 as necessary for multiple pages
# there are many more payload options available for filtering such as #of bedrooms & baths, inspect the chrome request for all options
payload = {
    "ZoomLevel": "11",
    "LatitudeMax": LATMAX,
    "LongitudeMax": LONGMAX,
    "LatitudeMin": LATMIN,
    "LongitudeMin": LONGMIN,
    "Sort":"6-D",
    "PropertyTypeGroupID":"1",
    "PropertySearchTypeId":"1",
    "TransactionTypeId":"2",
    "Currency":"CAD",
    "RecordsPerPage":"12",
    "ApplicationId":"1",
    "CultureId":"1",
    "Version":"7.0",
    "CurrentPage": PAGE,  
}

# this is the same header information used by chrome when making the same rest request
header = {
    'authority': 'api2.realtor.ca',
    'accept': '*/*',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36',
    'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'origin': 'https://www.realtor.ca',
    'sec-fetch-site': 'same-site',
    'sec-fetch-mode': 'cors',
    'sec-fetch-dest': 'empty',
    'referer': 'https://www.realtor.ca/',
    'accept-language': 'en-US,en;q=0.9',
}


req = requests.post('https://api2.realtor.ca/Listing.svc/PropertySearch_Post', headers=header, cookies=req.cookies, data=payload)


# the result from the API is a json file of 12 property listings, see response.json for example response
print(json.dumps(json.loads(req.content)))
