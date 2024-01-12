import requests

url = 'https://test.manage.test.iko.max-it-eam.com/maximo/api/os/MXAPIWODETAIL?lean=1&oslc.where=wonum="W14F10038" and siteid="KLU"&oslc.select=*'

headers = {
  'apikey': 'smqtrblh2ofemdkgl0roh1vbod7k6ph3ua9kgg3h',
}

response = requests.get(url, headers=headers)

wo_id = response.json()['member'][0]['href'].split('/')[-1]

print(response.text)

url = "https://test.manage.test.iko.max-it-eam.com/maximo/api/os/MXAPIWODETAIL/xxxx/doclinks"

url = url.replace('xxxx', wo_id)

with open("C:\\Users\\majona\\Documents\\Pirana\\photos\\0AF5140D-6484-4E48-A40E-5EA721B2E080.jpg", 'rb') as f:
    payload = f.read()

headers = {
  'apikey': 'smqtrblh2ofemdkgl0roh1vbod7k6ph3ua9kgg3h',
  'slug': 'testFile.jpg',
  'X-document-meta': 'Attachments',
  # 'Content-Type': 'image/jpeg',
}

response = requests.post(url, headers=headers, data=payload)

print(response.text)
