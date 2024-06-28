import requests
import csv

apiKey = 'smqtrblh2ofemdkgl0roh1vbod7k6ph3ua9kgg3h'
headers = {'apikey': apiKey}

with open('Python/AssetUpload/hits.csv', "r") as f:
    reader = csv.reader(f)
    data = list(reader)

for row in data:
    response = requests.delete(url=row[0], headers=headers)
    print(response.status_code)