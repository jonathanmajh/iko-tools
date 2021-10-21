import re
import csv

manus = {}
with open('ItemDescriptions/companies.txt',  encoding='utf8') as f:
    for line in f:
        manus[line.strip()] = 1

items = {}
with open('ItemDescriptions/ItemDescriptions.txt',  encoding='utf8') as f:
    for line in f:
        description = line.strip().split(',')
        item_type = description.pop(0)
        if not(item_type in items):
            items[item_type] = {}
        for word in description:
            match = re.search('[0-9]+', word)
            if not match and not (word in manus):
                if not(word in items[item_type]):
                    items[item_type][word] = 1
                else:
                    items[item_type][word] = items[item_type][word] + 1

with open('ItemDescriptions/result.csv', 'w', encoding='UTF8', newline='') as f:
    write = csv.writer(f)
    for item in items:
        for word in items[item]:
            row = [item, word, items[item][word], f'{item}{items[item][word]}']
            write.writerow(row)
# write as {item_type}{count} < this should be sortable while keeping item_type grouped
#print(items)