import os
import glob
import re
from datetime import datetime, timedelta, time

import fitz
from openpyxl import Workbook

from dataclasses import dataclass

# ------ CHANGE ME ---------
folder_path = "C:\\Users\\majona\\GitHub\\iko-tools\\Python\\antwerpTimeConverter\\input"
save_path = f'C:\\Users\\majona\\GitHub\\iko-tools\\Python\\antwerpTimeConverter\\{datetime.now().strftime("%d-%m-%Y_%H-%M-%S")}.xlsx'
year = '2023'


@dataclass
class person:
    include: bool = False
    laborcode: str = "blank"

# add more Nummer to personid codes as necessary
people = {
    "21592": person(True, "BEANJBRU"),
    "21616": person(True, "BEANDCAR"),
    "21612": person(True, "BEANDHOS"),
    "21545": person(True, "BEANPJOR"),
    "21263": person(True, "BEANBMAE"),
    "21368": person(True, "BEANPMAE"),
    "21457": person(True, "BEANAMAR"),
    "21547": person(True, "BEANNPAT"),
    "20746": person(True, "BEANDROO"),
    "21328": person(True, "BEANPVHA"),
    "20889": person(True, "BEANBVRI"),
    "21193": person(True, "BEANBWIN"),
}

converted_wb = Workbook()
converted_ws = converted_wb.create_sheet("Sheet1", 0)
converted_ws.append(
    [
        "Empl#",
        "Employee Name",
        "Day",
        "Time In",
        "Time Out",
        "Hours",
        "LL1",
        "LL2",
        "LL3",
        "LL4",
        "LL5",
        "LL6",
        "LL7",
        "Reg",
        "O/T",
    ]
)

date_regex = r"^\d{2}/\d{2}/\d{4}"
time_regex = r"\b\d{1,2}:\d{2}\b"

for filepath in glob.glob(os.path.join(folder_path, "*.pdf")):
    doc = fitz.open(filepath)
    print(f'Reading: {filepath}')
    employee = person()
    for page in doc:
        # the blocks option seems to give best result where lines are kept together
        text = page.get_text('blocks')
        # print(text)
        for row in text:
            # see if employee is defined if the row is a employee definition
            if (row[4].find('Nummer') > 0):
                person_num = row[4][row[4].find('Nummer') + 7:row[4].find('Badge')]
                # print(person_num)
                employee = people.get(person_num.strip())
                if (not(employee.include)):
                    print(f'Not Found: {person_num.strip()}')
                continue
            if (employee.include):
                match = re.findall(date_regex, row[4])
                if (match):
                    # print(match)
                    time_matches = re.findall(time_regex, row[4])
                    hours = row[4][row[4].find(time_matches[1])+len(time_matches[1]):].split('\n')
                    # print(hours)
                    start_time = datetime.strptime(f'{match[0]} {time_matches[0]}', "%d/%m/%Y %H:%M")
                    end_time = start_time + timedelta(hours=float(hours[1].replace(',','.')))
                    converted_ws.append(
                        [
                            person_num.strip(),
                            employee.laborcode,
                            start_time.strftime('%A').upper(),
                            start_time.strftime('%Y-%m-%d %H:%M'),
                            end_time.strftime('%Y-%m-%d %H:%M'),
                            hours[1].replace(',','.'),
                            None,
                            None,
                            None,
                            None,
                            None,
                            None,
                            None,
                            hours[1].replace(',','.'),
                            0,
                        ]
                    )
                else:
                    print(f'Skipping row: {row[4]}')
                    continue
                
converted_wb.save(save_path)