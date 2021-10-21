import csv
from os import read

jps = {}

with open('JobTasksToDescription/BL_Route_JobTasks.csv') as f:
    reader = csv.reader(f, delimiter=',')
    for row in reader:
        if (not(row[0] in jps)):
            jps[row[0]] = f'{row[1]} - {row[5]} - {row[6]}\n'
        else:
            jps[row[0]] = f'{jps[row[0]]}{row[1]} - {row[5]} - {row[6]}\n'

with open('JobTasksToDescription/result.csv', 'w', encoding='UTF8', newline='') as f:
    write = csv.writer(f)
    for item in jps:
        row = [item, jps[item]]
        write.writerow(row)