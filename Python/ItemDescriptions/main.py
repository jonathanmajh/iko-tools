import re

# TODO remove manufacturers
# TODO how print :thinking:

items = {}
with open('ItemDescriptions/ItemDescriptions.txt',  encoding='utf8') as f:
    for line in f:
        description = line.strip().split(',')
        item_type = description.pop(0)
        if not(item_type in items):
            items[item_type] = {}
        for word in description:
            match = re.search('[0-9]+', word)
            if not match:
                if not(word in items[item_type]):
                    items[item_type][word] = 1
                else:
                    items[item_type][word] = items[item_type][word] + 1

# write as {item_type}{count} < this should be sortable while keeping item_type grouped
print(items)