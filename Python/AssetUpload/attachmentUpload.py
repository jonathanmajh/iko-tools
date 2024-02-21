import csv
import requests

headers = {
  'apikey': 'smqtrblh2ofemdkgl0roh1vbod7k6ph3ua9kgg3h',
}

with open('Python\\AssetUpload\\KlundertWOImages.csv') as csv_file:
    csv_reader = csv.reader(csv_file, delimiter=',')
    for row in csv_reader:
        url = f'https://prod.manage.prod.iko.max-it-eam.com/maximo/api/os/MXAPIWODETAIL?lean=1&oslc.where=wonum="{row[1]}" and siteid="KLU"&oslc.select=*'
        response = requests.get(url, headers=headers)

        wo_id = response.json()['member'][0]['href'].split('/')[-1]

        # print(response.text)
        print(wo_id)

        url = f"https://prod.manage.prod.iko.max-it-eam.com/maximo/api/os/MXAPIWODETAIL/{wo_id}/doclinks"

        with open(f"C:\\Users\\majona\\Documents\\Code\\iko-tools\\Python\\SQL\\photos\\{row[0]}.jpg", 'rb') as f:
            payload = f.read()

        headers = {
          'apikey': 'smqtrblh2ofemdkgl0roh1vbod7k6ph3ua9kgg3h',
          'slug': f'{row[1]}.jpg',
          'X-document-meta': 'Attachments',
          # 'Content-Type': 'image/jpeg',
        }

        response = requests.post(url, headers=headers, data=payload)

        print(response.text)
