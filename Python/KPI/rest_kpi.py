import csv
from ctypes import alignment
from textwrap import fill
import requests
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import NamedStyle, Font, Border, Side, PatternFill, Alignment
from openpyxl.formatting.rule import IconSetRule

#REMEMBER TO MOVE LABOR CHARGED TO LOWEST LEVEL CHILD ASSET TO THE END
# --------------------------------
# variables to change
month = 12
year = 2022
# variables to change
# --------------------------------
current_month = year * 12 + month

ENDPOINTS = {
    'IKO_KPI_HVHCRITICALASSETPMCOMPLETEDONTIME': 'VH & H Critical Asset PM Completed On Time Trend',
    'IKO_KPI_MAXIMOVSKRONOSHOURS': 'Maximo Lbr Hrs vs Kronos Lb Hrs',
    'IKO_KPIWOW1STWHY': "Work Orders with 5 Why's",
    'IKO_KPI_LABORHOURCHARGEDTOLLCA':
    'Labor Hrs Charged to Lowest Child Assets',
    'IKO_KPI_MROLINEFROMMAXIMO': 'MRO PR Lines from Maximo',
    'IKO_API_ASSETSPAREPARTCOUNTWITHCRITICALITY': 'Asset Spare Parts Count',
    'IKO_KPI_ISSUEDITEMLISTEDINSPAREPARTLIST':
    'Issued Items Listed in Spare Part List',
    'IKO_KPI_CYCLECOUNT': 'Cycle Count',
    'IKO_KPI_SCHEDULEDWORKORDERS': 'Scheduled Work Orders',
    'IKO_KPI_WONOTCREATEDBYPLANNER': 'WO Not Created by Planner',
    'IKO_KPI_JOBPLANCREATED': 'Job Plans Created/Updated',
}

averages = {
    'IKO_KPI_HVHCRITICALASSETPMCOMPLETEDONTIME': [0, 0],
    'IKO_KPI_MAXIMOVSKRONOSHOURS': [0, 0],
    'IKO_KPIWOW1STWHY': [0, 0],
    'IKO_KPI_LABORHOURCHARGEDTOLLCA': [0, 0],
    'IKO_KPI_MROLINEFROMMAXIMO': [0, 0],
    'IKO_API_ASSETSPAREPARTCOUNTWITHCRITICALITY': [0, 0],
    'IKO_KPI_ISSUEDITEMLISTEDINSPAREPARTLIST': [0, 0],
    'IKO_KPI_CYCLECOUNT': [0, 0],
    'IKO_KPI_SCHEDULEDWORKORDERS': [0, 0],
    'IKO_KPI_WONOTCREATEDBYPLANNER': [0, 0],
    'IKO_KPI_JOBPLANCREATED': [0, 0],
}

overall = {'up': 0, 'side': 0, 'down': 0}

ENDPOINT_URL = list(ENDPOINTS)

PMONTIME = 'IKO_KPI_PMCOMPLETEDONTIME'

SITES = {
    # 'GI': 'Madoc',
    # 'GE': 'Ashcroft',
    # 'GS': 'Sylacauga',
    # 'BA': 'Calgary',
    # 'GV': 'Hillsboro',
    # 'GH': 'Hawkesbury',
    # 'AA': 'IKO Brampton',
    # 'GJ': 'CRC Toronto',
    'CA': 'Kankakee',
    # 'GC': 'Sumas',
    # 'GK': 'IG Brampton',
    # 'CAM': 'Appley Bridge',
    # 'BL': 'Hagerstown',
    # 'RAM': 'Alconbury',
    # 'GP': 'CRC Brampton',
    # 'GM': 'IG High River',
    # 'COM': 'Combronde',
    # 'GR': 'Bramcal',
    # 'GX': 'MaxiMix'
}

MONTHS = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
]

URL = [
    'http://nscandacmaxapp1.na.iko/maximo/oslc/script/',
    '?site=',
    '&_lid=majona&_lpwd=happy818',
]

results = [
    ["Site"],
    ["KPI metric"],
    [],  # spacer
    ['Average'],
    [],  # spacer
    [MONTHS[month - 1]],
    [MONTHS[month - 2]],
    [MONTHS[month - 3]],
    [MONTHS[month - 4]],
    [MONTHS[month - 5]],
    [MONTHS[month - 6]],
    [MONTHS[month - 7]],
    [MONTHS[month - 8]],
    [MONTHS[month - 9]],
    [MONTHS[month - 10]],
    [MONTHS[month - 11]],
    [MONTHS[month % 12]],
]


def request_wrapper(url):
    r = requests.get(url)
    if (r.status_code != 200):
        print(f'Error: {r.status_code}')
        print(f'Error: {url}')
        result = {'target': -1}
        return result 
    else:
        r = r.json()
        if (not 'info' in r):
            r['info'] = []
        return r


def level(table):
    columns = len(table[0])
    for j in range(len(table)):
        while (len(table[j]) < columns):
            table[j].append('-')
    return table


def zero(table):
    columns = len(table[0])
    for j in range(len(table)):
        while (len(table[j]) < columns):
            table[j].append(0)
    return table


def compare_values(target, results, ws, ws_write_col):
    # shouldnt need to return ws back
    if (results[3][-1] != '-' and target != '-'):
        if ((target <= results[3][-1])):
            ws.cell(row=3 * site_i + 1, column=ws_write_col).fill = green
        elif ((target - results[3][-1]) < (0.2 * target)):
            ws.cell(row=3 * site_i + 1, column=ws_write_col).fill = yellow
        else:
            ws.cell(row=3 * site_i + 1, column=ws_write_col).fill = red
        ws.cell(row=3 * site_i + 1, column=ws_write_col).number_format = '0%'
    if (results[5][-1] != '-'):
        if ((target <= results[5][-1])):
            ws.cell(row=3 * site_i + 2, column=ws_write_col).fill = green
        elif ((target - results[5][-1]) < (0.2 * target)):
            ws.cell(row=3 * site_i + 2, column=ws_write_col).fill = yellow
        else:
            ws.cell(row=3 * site_i + 2, column=ws_write_col).fill = red
        ws.cell(row=3 * site_i + 2, column=ws_write_col).number_format = '0%'
        # record value for overall average
        averages[ENDPOINT_URL[ws_write_col - 8]][0] += results[5][-1]
        averages[ENDPOINT_URL[ws_write_col - 8]][1] += 1
    if ((results[5][-1] == '-') or (results[3][-1] == '-')):  # blank is meh
        ws.cell(row=3 * site_i + 3, column=ws_write_col, value=0)
        overall['side'] += 1
    elif (abs(results[3][-1] - results[5][-1]) <
          0.005):  # very small change is meh
        ws.cell(row=3 * site_i + 3, column=ws_write_col, value=0)
        overall['side'] += 1
    elif (results[3][-1] < results[5][-1]):  # above average
        ws.cell(row=3 * site_i + 3, column=ws_write_col, value=1)
        overall['up'] += 1
    else:  # below average
        ws.cell(row=3 * site_i + 3, column=ws_write_col, value=-1)
        overall['down'] += 1
    ws.cell(row=3 * site_i + 1, column=ws_write_col).alignment = center
    ws.cell(row=3 * site_i + 2, column=ws_write_col).alignment = center
    ws.cell(row=3 * site_i + 3, column=ws_write_col).alignment = center


def generic_result_reader(url, results, ws, ws_write_col):
    result = request_wrapper(url)
    if (result == -1):
        for j in range(len(results) - 2):
            results[j + 2].append("-")
    else:
        overall_percent = 0.0
        for item in result['info']:
            monthdate = item['year'] * 12 + item['month']
            results[5 + (current_month - monthdate)].append(item['percentage'])
            overall_percent += item['percentage']
        if (len(result['info']) > 0):
            results[3].append(overall_percent / len(result['info']))
    results = level(results)
    ws.cell(row=3 * site_i + 1, column=ws_write_col, value=results[3][-1])
    ws.cell(row=3 * site_i + 2, column=ws_write_col, value=results[5][-1])
    compare_values(result['target'], results, ws, ws_write_col)


wb = Workbook()
ws = wb.active
site_i = 0

#text alignment
center = Alignment(horizontal='center', vertical='center')
ninety_center = Alignment(horizontal='center',
                          vertical='center',
                          text_rotation=90)

#cell colors
green = PatternFill(start_color='00C6EFCE',
                    end_color='00C6EFCE',
                    fill_type='solid')
yellow = PatternFill(start_color='00FFEB9C',
                     end_color='00FFEB9C',
                     fill_type='solid')
red = PatternFill(start_color='00FFC7CE',
                  end_color='00FFC7CE',
                  fill_type='solid')
white = PatternFill(fill_type=None,
                    start_color='FFFFFFFF',
                    end_color='FF000000')
#border styles
thin = Side(border_style="thin", color="000000")
double = Side(border_style="medium", color="000000")

# format workbook
ws.append([])
# should be generated from above list
ws.append([
    '',
    '',
    '',
    '',
    '',
    'KPI Summary',
    '',
    'VH & H Critical Asset PM Completed On Time Trend',
    'Maximo Lbr Hrs vs Kronos Lb Hrs',
    "Work Orders with 5 Why's",
    'Issued Items Listed in Spare Part List',
    'MRO PR Lines from Maximo',
    'Asset Spare Parts Count',
    'Cycle Count',
    'Scheduled Work Orders',
    'WO Not Created by Planner',
    'Job Plans Created/Updated',
    'Labor Hrs Charged to Lowest Child Assets',
])
ws.merge_cells(start_row=2, start_column=6, end_row=2, end_column=7)
for cell in range(13):
    ws.cell(row=2, column=6 + cell).alignment = Alignment(text_rotation=45)
ws.append([
    '',
    '',
    '',
    '',
    '',
    'Target',
    '',
    '95%',
    '80%',
    '50%',
    '80%',
    '80%',
    '-',
    '10%',
    '80%',
    '90%',
    25,
    '80%',
])
ws.merge_cells(start_row=3, start_column=6, end_row=3, end_column=7)
for cell in range(13):
    ws.cell(row=3, column=6 + cell).border = Border(bottom=double, top=double)
    ws.cell(row=2, column=6 + cell).border = Border(top=double)
    ws.cell(row=3, column=6 + cell).alignment = center

ws.cell(row=4, column=2, value='GRANULES')
ws.cell(row=4, column=2).alignment = ninety_center
ws.cell(row=4, column=2).border = Border(top=double,
                                         left=double,
                                         right=double,
                                         bottom=double)
ws.cell(row=4, column=2).font = Font(name='Calibri', size=12, bold=True)
ws.merge_cells(start_row=4, start_column=2, end_row=9, end_column=2)
ws.column_dimensions['B'].width = 2
ws.cell(row=10, column=3, value='SHINGLE')
ws.cell(row=10, column=3).alignment = ninety_center
ws.cell(row=10, column=3).border = Border(top=double,
                                          left=double,
                                          right=double,
                                          bottom=double)
ws.cell(row=10, column=3).font = Font(name='Calibri', size=12, bold=True)
ws.merge_cells(start_row=10, start_column=3, end_row=33, end_column=3)
ws.column_dimensions['C'].width = 2
ws.cell(row=28, column=4, value='MEMBRANES')
ws.cell(row=28, column=4).alignment = ninety_center
ws.cell(row=28, column=4).border = Border(top=double,
                                          left=double,
                                          right=double,
                                          bottom=double)
ws.cell(row=28, column=4).font = Font(name='Calibri', size=12, bold=True)
ws.merge_cells(start_row=28, start_column=4, end_row=42, end_column=4)
ws.column_dimensions['D'].width = 2
ws.cell(row=40, column=5, value='ISO')
ws.cell(row=40, column=5).alignment = ninety_center
ws.cell(row=40, column=5).border = Border(top=double,
                                          left=double,
                                          right=double,
                                          bottom=double)
ws.cell(row=40, column=5).font = Font(name='Calibri', size=12, bold=True)
ws.merge_cells(start_row=40, start_column=5, end_row=54, end_column=5)
ws.column_dimensions['E'].width = 2
ws.cell(row=2, column=6).border = Border(left=double,
                                         bottom=double,
                                         top=double)
ws.cell(row=3, column=6).border = Border(left=double,
                                         bottom=double,
                                         top=double)
ws.cell(row=2, column=18).border = Border(right=double,
                                          bottom=double,
                                          top=double)
ws.cell(row=3, column=18).border = Border(right=double,
                                          bottom=double,
                                          top=double)
ws.column_dimensions['F'].width = 13
# iconset test
#rule = IconSetRule('5Arrows', 'num', [-2, -1, 0, 1, 2], showValue=False, percent=None, reverse=None)

for site in SITES:
    ws_write_col = 8
    site_i += 1
    ws.cell(row=3 * site_i + 2, column=6, value=SITES[site])
    ws.cell(row=3 * site_i + 1, column=6).alignment = center
    ws.cell(row=3 * site_i + 1, column=6).border = Border(left=double,
                                                          top=double)
    ws.cell(row=3 * site_i + 2, column=6).border = Border(left=double)
    ws.cell(row=3 * site_i + 3, column=6).border = Border(left=double)
    ws.cell(row=3 * site_i + 1, column=7, value='Avg')
    ws.cell(row=3 * site_i + 1, column=7).alignment = center
    ws.cell(row=3 * site_i + 2, column=7, value=MONTHS[month - 1])
    ws.cell(row=3 * site_i + 2, column=7).alignment = center
    ws.cell(row=3 * site_i + 3, column=7, value='Trend')
    ws.cell(row=3 * site_i + 3, column=7).alignment = center
    for cell in range(12):
        ws.cell(row=3 * site_i + 3,
                column=7 + cell).border = Border(bottom=double)
    ws.cell(row=3 * site_i + 1, column=18).border = Border(right=double,
                                                           top=double)
    ws.cell(row=3 * site_i + 2, column=18).border = Border(right=double)
    ws.cell(row=3 * site_i + 3, column=18).border = Border(right=double)

    # pm on time
    results[0].append(site)
    results[1].append(ENDPOINT_URL[0])
    print(f'{ENDPOINT_URL[0]} : {site}')
    result = request_wrapper(
        f'{URL[0]}{ENDPOINT_URL[0]}{URL[1]}{site}{URL[2]}')
    if (result == -1):
        for j in range(len(results) - 2):
            results[j + 2].append("-")
    else:
        allcount = {}
        ontimecount = {}
        percentage = {}
        overall_percent = 0.0
        for item in result['info']:
            monthdate = item['year'] * 12 + item['month']
            if (monthdate in allcount):
                allcount[monthdate] += item['allcount']
                ontimecount[monthdate] += item['ontimecount']
            else:
                allcount[monthdate] = item['allcount']
                ontimecount[monthdate] = item['ontimecount']
        for key in allcount.keys():
            results[5 + (current_month - key)].append(ontimecount[key] /
                                                      allcount[key])
            overall_percent += ontimecount[key] / allcount[key]
        if (len(allcount.keys()) > 0):
            results[3].append(overall_percent / len(allcount.keys()))
        else:
            print('Falling back to All PM Completed on Time KPI')
            result = request_wrapper(
                f'{URL[0]}{PMONTIME}{URL[1]}{site}{URL[2]}')
            if (result != -1):
                for item in result['info']:
                    monthdate = item['year'] * 12 + item['month']
                    results[5 + (current_month - monthdate)].append(
                        item['percentage'])
                    overall_percent += item['percentage']
                if (len(result['info']) > 0 and overall_percent > 0.001):
                    results[3].append(overall_percent / len(result['info']))
    results = level(results)
    ws.cell(row=3 * site_i + 1, column=ws_write_col, value=results[3][-1])
    ws.cell(row=3 * site_i + 2, column=ws_write_col, value=results[5][-1])
    compare_values(result['target'], results, ws, ws_write_col)

    ws_write_col += 1
    # kronos vs maximo
    results[0].append(site)
    results[1].append(ENDPOINT_URL[1])
    print(f'{ENDPOINT_URL[1]} : {site}')
    generic_result_reader(f'{URL[0]}{ENDPOINT_URL[1]}{URL[1]}{site}{URL[2]}',
                          results, ws, ws_write_col)
    
    ws_write_col += 1
    # WO With Why
    results[0].append(site)
    results[1].append(ENDPOINT_URL[2])
    print(f'{ENDPOINT_URL[2]} : {site}')
    generic_result_reader(f'{URL[0]}{ENDPOINT_URL[2]}{URL[1]}{site}{URL[2]}',
                          results, ws, ws_write_col)
    
    ws_write_col += 1
    # IKO_KPI_ISSUEDITEMLISTEDINSPAREPARTLIST
    results[0].append(site)
    results[1].append(ENDPOINT_URL[6])
    print(f'{ENDPOINT_URL[6]} : {site}')
    generic_result_reader(f'{URL[0]}{ENDPOINT_URL[6]}{URL[1]}{site}{URL[2]}',
                          results, ws, ws_write_col)

    ws_write_col += 1
    # IKO_KPI_MROLINEFROMMAXIMO
    results[0].append(site)
    results[1].append(ENDPOINT_URL[4])
    print(f'{ENDPOINT_URL[4]} : {site}')
    generic_result_reader(f'{URL[0]}{ENDPOINT_URL[4]}{URL[1]}{site}{URL[2]}',
                          results, ws, ws_write_col)

    ws_write_col += 1
    # IKO_API_ASSETSPAREPARTCOUNTWITHCRITICALITY
    results[0].append(site)
    results[1].append(ENDPOINT_URL[5])
    print(f'{ENDPOINT_URL[5]} : {site}')
    result = request_wrapper(
        f'{URL[0]}{ENDPOINT_URL[5]}{URL[1]}{site}{URL[2]}')
    if (result == -1):
        for j in range(len(results) - 2):
            results[j + 2].append("-")
    else:
        allcount = {}
        overall_percent = 0.0
        for item in result['info']:
            monthdate = item['year'] * 12 + item['month']
            overall_percent += item['sparepartcount']
            if (monthdate in allcount):
                allcount[monthdate] += item['sparepartcount']
            else:
                allcount[monthdate] = item['sparepartcount']
        for monthdate in allcount.keys():
            results[5 + (current_month - monthdate)].append(
                allcount[monthdate])
        if (len(result['info']) > 0):
            results[3].append(overall_percent / len(allcount.keys()))
    results = level(results)
    ws.cell(row=3 * site_i + 1, column=ws_write_col, value=results[3][-1])
    ws.cell(row=3 * site_i + 2, column=ws_write_col, value=results[5][-1])
    compare_values(result['target'], results, ws, ws_write_col)
    ws.cell(row=3 * site_i + 1, column=ws_write_col).number_format = '#,##0'
    ws.cell(row=3 * site_i + 2, column=ws_write_col).number_format = '#,##0'
    ws.cell(row=3 * site_i + 1, column=ws_write_col).fill = white
    ws.cell(row=3 * site_i + 2, column=ws_write_col).fill = white

    ws_write_col += 1
    # IKO_KPI_CYCLECOUNT
    results[0].append(site)
    results[1].append(ENDPOINT_URL[7])
    print(f'{ENDPOINT_URL[7]} : {site}')
    generic_result_reader(f'{URL[0]}{ENDPOINT_URL[7]}{URL[1]}{site}{URL[2]}',
                          results, ws, ws_write_col)

    ws_write_col += 1
    # IKO_KPI_SCHEDULEDWORKORDERS
    results[0].append(site)
    results[1].append(ENDPOINT_URL[8])
    print(f'{ENDPOINT_URL[8]} : {site}')
    generic_result_reader(f'{URL[0]}{ENDPOINT_URL[8]}{URL[1]}{site}{URL[2]}',
                          results, ws, ws_write_col)

    ws_write_col += 1
    # IKO_KPI_WONOTCREATEDBYPLANNER
    results[0].append(site)
    results[1].append(ENDPOINT_URL[9])
    print(f'{ENDPOINT_URL[9]} : {site}')
    generic_result_reader(f'{URL[0]}{ENDPOINT_URL[9]}{URL[1]}{site}{URL[2]}',
                          results, ws, ws_write_col)

    ws_write_col += 1
    # IKO_KPI_JOBPLANCREATED
    results[0].append(site)
    results[1].append(ENDPOINT_URL[10])
    print(f'{ENDPOINT_URL[10]} : {site}')
    result = request_wrapper(
        f'{URL[0]}{ENDPOINT_URL[10]}{URL[1]}{site}{URL[2]}')
    if (result == -1):
        for j in range(len(results) - 2):
            results[j + 2].append("-")
    else:
        overall_percent = 0.0
        for item in result['info']:
            monthdate = item['year'] * 12 + item['month']
            results[5 + (current_month - monthdate)].append(item['total'])
            overall_percent += item['total']
        if (len(result['info']) > 0):
            results[3].append(overall_percent / len(result['info']))
    results = level(results)
    ws.cell(row=3 * site_i + 1, column=ws_write_col, value=results[3][-1])
    ws.cell(row=3 * site_i + 2, column=ws_write_col, value=results[5][-1])
    compare_values(result['target'], results, ws, ws_write_col)
    ws.cell(row=3 * site_i + 1, column=ws_write_col).number_format = '#,##0'
    ws.cell(row=3 * site_i + 2, column=ws_write_col).number_format = '#,##0'

    ws_write_col += 1
    # IKO_KPI_LABORHOURCHARGEDTOLLCA
    results[0].append(site)
    results[1].append(ENDPOINT_URL[3])
    print(f'{ENDPOINT_URL[3]} : {site}')
    generic_result_reader(f'{URL[0]}{ENDPOINT_URL[3]}{URL[1]}{site}{URL[2]}',
                          results, ws, ws_write_col)

for cell in range(13):
    ws.cell(row=61, column=6 + cell).border = Border(top=double)

i = 0
for average in averages:
    ws.cell(row=70,
            column=8 + i,
            value=averages[average][0] / averages[average][1])
    i += 1

ws.cell(row=72, column=6, value='Up')
ws.cell(row=72, column=7, value=overall['up'])
ws.cell(row=73, column=6, value='Same')
ws.cell(row=73, column=7, value=overall['side'])
ws.cell(row=74, column=6, value='Down')
ws.cell(row=74, column=7, value=overall['down'])

# write
with open(
        f"C:\\Users\\majona\\GitHub\\iko-tools\\Python\\KPI\\{date.today().isoformat()}_KPI.csv",
        "w",
        newline="",
) as f:
    writer = csv.writer(f)
    writer.writerows(results)

wb.save(
    filename=
    f"C:\\Users\\majona\\GitHub\\iko-tools\\Python\\KPI\\{date.today().isoformat()}_KPI.xlsx"
)
