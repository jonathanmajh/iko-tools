# will need rework RE integers as strings before using

import requests
import csv
import logging

logging.basicConfig(filename='log.log',
                            filemode='a',
                            format='%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s',
                            datefmt='%H:%M:%S',
                            level=logging.INFO)

logger = logging.getLogger('urbanGUI')

jp_mats = {}

with open('C:\\Users\\majona\\Documents\\GitHub\\iko-tools\\Python\\jobmaterial\\GI_JobMaterials.csv', newline='') as f:
    reader = csv.reader(f)
    data = list(reader)

f = open('C:\\Users\\majona\\Documents\\GitHub\\iko-tools\\Python\\jobmaterial\\add_to_mats.csv', 'w')
writer = csv.writer(f)

for row in data:
    if jp_mats.get(row[2]):
        jp_mats[row[2]].append([row[0],row[1]])
    else:
        jp_mats[row[2]] = [[row[0],row[1]]]

for jpnum in jp_mats: 
    response = requests.get(f'http://nscandacmaxapp1/maxrest/oslc/os/iko_jobmaterial?oslc.select=*&oslc.where=jpnum="{jpnum}"&_lid=corcoop3&_lpwd=happy818')
    response = response.json()
    logging.info(jpnum)
    try:
        response["rdfs:member"][0]["spi:jobmaterial"]
    except:
        pass
    else:
        for jobmat in response["rdfs:member"][0]["spi:jobmaterial"]:
            item_qty = [jobmat['spi:itemnum'], jobmat['spi:itemqty']]
            if item_qty in jp_mats[jpnum]:
                jp_mats[jpnum].remove(item_qty)
                logging.info(f'keep:   {item_qty}')
            else:
                logging.info(f'remove: {item_qty}')
                response2 = requests.delete(f"{jobmat['localref'].replace('localhost', 'nscandacmaxapp1')}?_lid=corcoop3&_lpwd=happy818")
                logging.info(response2)
    for item in jp_mats[jpnum]:
        logging.info(f'add:    {item}')
        temp = ['IKO-CAD','GI',jpnum,0,item[0],item[1],'ITEMSET1']
        writer.writerow(temp)

f.close()