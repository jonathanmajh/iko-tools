import requests

req = requests.get('https://www.realtor.ca/')

cookies = list(req.cookies)

payload = {
    "ZoomLevel": "11",
    "LatitudeMax": "43.90135",
    "LongitudeMax": "-79.05332",
    "LatitudeMin": "43.51420",
    "LongitudeMin": "-79.69945",
    "Sort":"6-D",
    "PropertyTypeGroupID":"1",
    "PropertySearchTypeId":"1",
    "TransactionTypeId":"2",
    "Currency":"CAD"
    "RecordsPerPage":"12",
    "ApplicationId":"1",
    "CultureId":"1",
    "Version":"7.0",
    "CurrentPage":"1",  
}

cookies = {
    "visid_incap_2269415":"fC6X4ONfQa+zVdsZD7s7g+/Yz18AAAAAQUIPAAAAAAA6Tm2BkG9fKUHKN8+tvH3O"
    "nlbi_2269415":"Wu5QMjzf7UAqlsvVAfQG9AAAAABlMHyNWOMHFt+2/ADvdQ+y",
    "incap_ses_305_2269415":"M0jFHIuickH5T31+KpQ7BO/Yz18AAAAAW/VyEgfWzppSoMyz3I1f2A==",
    "gig_bootstrap_3_mrQiIl6ov44s2X3j6NGWVZ9SDDtplqV7WgdcyEpGYnYxl7ygDWPQHqQqtpSiUfko":"gigya-pr_ver3",
    "ASP.NET_SessionId":"mgkv2ttopgh0r5clgx4hxfsh",
    "visid_incap_2271082":"DhLQP7TqS6eSe6TKrbjOwITZz18AAAAAQUIPAAAAAABW5DUsJ31rJWGO42XUCW76",
    "nlbi_2271082":"ioRDJ4AnexHJ0ap3/WCXVgAAAADwjO/iuokgC86y5MKhzNfO",
    "incap_ses_305_2271082":"BzZUHB9PzwxdCX5+KpQ7BITZz18AAAAAMu6w/tHI43lg1lRToZdcnQ==",
    "nlbi_2269415_2147483646":"xG3ccAWWbBv+PiaNAfQG9AAAAADq9C5NCaYjVe4ol1VBwYXH",
    "reese84":"3:VLX2WjIazRlNweS5FcYl1Q==:QBPREpAj9s6U3Rl74GnYvGSEzANp/5/t63A4lhZOEc0mgDmV2G5ZXnRpoh2Zu1FK0pVMxsyBNu66yVMIQjO9shce1dmCrC4Ux0SArKi6xEpFZdA3PGOHILHtetVyKZAOJisnmbVgWEJo4HwuT4zMMvobWlg8jfNWcDUfUm+w09P2nEkReHYF3KOWBth4WXKWTFYq79OCKqUzqvt7pS4uTHeyF5Mxb/zaLadON1d++OIgt/h/SnyyrjjW1JzSlPenAhRtY9sFNGiFWXgMQlOH6lpg2oyWWm6e7/PxVftKzRE9ek/f4ll93CCz7GalF0+Ds3qP8bIic6s22acCvST3iyPjz4uk1++9WOlrJQpxJLSS6+eNJYxeKYKOqQ1PlSLAi9MS/KaPwfAOPVXzwTOE+qp6QoJArTowjYU3WivbIO8=:6DOPOMyGvnCW9nIoXWGS+SHGY0kQNF2Mbwg339MxDic="
}

print('end')