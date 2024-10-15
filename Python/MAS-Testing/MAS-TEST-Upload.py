import requests

uploadFiles = [
    ['IKO_ITEMMASTER','IKO_ITEMMASTER.csv',],
    ['IKO_INVENTORY','IKO_INVENTORY.csv',],
    ['IKO_LOCATION','IKO_LOCATION.csv',],
    ['IKO_ASSET','IKO_ASSET.csv',],
    ['IKO_JOBPLAN','IKO_JOBPLAN.csv',],
    ['IKO_JPASSETLINK', 'IKO_JPASSETLINK.csv',],
    ['IKO_JOBLABOR', 'IKO_JOBLABOR.csv'],
]

uploadFiles2 = [
    ['IKO_PERSON','IKO_PERSON2.csv',],
    ['IKO_PERUSER','IKO_PERUSER.csv',],
    ['IKO_GROUPUSER','IKO_GROUPUSER.csv',],
    ['IKO_LABOR','IKO_LABOR.csv',],
]

baseUrl = 'https://test2.manage.test2.iko.max-it-eam.com/maximo/api/os/xxx?action=importfile&lean=1'
apiKey = 'k0do32u3h69bv0rd2sfoq1hqgflouetrpit3q4av'

def upload(uploadFiles):
    for file in uploadFiles:
        f = open(f'Python/MAS-Testing/{file[1]}', "r", encoding='utf-8-sig')
        body = f.read()
        url = baseUrl.replace('xxx', file[0])
        print('uploading ' + file[1])
        # preview
        headers = {'apikey': apiKey,'preview': '1','Content-Type': 'text/plain'}
        req = requests.post(url, headers=headers, data=body)
        print(req.text)
        # fr fr
        print('uploading fr fr ' + file[1])
        headers = {'apikey': apiKey,'Content-Type': 'text/plain'}
        req = requests.post(url, headers=headers, data=body)
        print(req.text)

# upload(uploadFiles)
# print('Waiting for Manual Item/Asset/Inventory/PM/User Creation/person groups...')
# input('Press Enter to Continue...')
# input('Are you sure?...')
upload(uploadFiles2)