import csv

input_file = 'Python/ReliabilityAlerts/SharepointReliabilityAlerts.csv'

asset_ids = []
asset_with_site = []
failure_classes = []
failure_classes_site = []

with open(input_file) as csv_file:
    csv_reader = csv.reader(csv_file, delimiter=',')
    line_count = 0
    for row in csv_reader:
        if line_count == 0:
            line_count += 1
        else:
            assets = row[3].split(';#')
            assets = [asset for asset in assets if len(asset)==5]
            [asset_ids.append(asset) for asset in assets]

            # get site id if it has been filled out
            site_id = row[2].split(':')[0]

            if site_id and assets:
                [asset_with_site.append(f'{site_id}_{asset}') for asset in assets]

            # process failure classes
            failure_class = row[9].split(';')
            failure_class = [fail.strip() for fail in failure_class if len(fail)>=6]
            [failure_classes.append(fail) for fail in failure_class]

            if site_id and failure_class:
                [failure_classes_site.append(f'{site_id}_{fail}') for fail in failure_class]

with open('ReliabilityAlerts.csv', 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerows([asset_with_site, failure_classes_site, asset_ids, failure_classes])
            