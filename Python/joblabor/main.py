import requests
import csv
import logging

logging.basicConfig(filename='log2.log',
                            filemode='a',
                            format='%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s',
                            datefmt='%H:%M:%S',
                            level=logging.INFO)

logger = logging.getLogger('urbanGUI')

with open('CheckJobLabor.csv', newline='') as f:
    reader = csv.reader(f)
    data = list(reader)

for row in data:
    response = requests.get(f'http://nscandacmaxapp1/maxrest/oslc/os/iko_joblabor?oslc.select=*&oslc.where=jpnum="{row[0]}"&_lid=corcoop3&_lpwd=happy818')
    response = response.json()
    labors = {}
    logging.info(row[0])
    for joblabor in response["rdfs:member"][0]["spi:joblabor"]:
        search = f"{joblabor['spi:orgid']}|{joblabor['spi:craft']}|{joblabor['spi:quantity']}|{joblabor['spi:laborhrs']}"
        exist = labors.get(search)
        if not exist:
            labors[search] = joblabor['localref']
            logging.info(f'adding   : {search}')
        else:
            logging.info(f'duplicate: {search}')
            response2 = requests.delete(f"{joblabor['localref'].replace('localhost', 'nscandacmaxapp1')}?_lid=corcoop3&_lpwd=happy818")
            logging.info(response2)