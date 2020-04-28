import json
import pprint
import webbrowser

import bs4
import requests
import xlwings

ota_url = 'http://192.168.10.6:8357'
mcu_url = 'http://192.168.10.1:5000/api/v1/dev/sn'
webbrowser.open(ota_url)
webbrowser.open(mcu_url)

response = requests.get(mcu_url)
file = open('sn.html', 'wb')
file.write(bytearray(response.text,encoding='utf-8'))
data = response.text
data = json.loads(data)['data']
# data={}
file = open('OTA.html', encoding='utf8')
soup = bs4.BeautifulSoup(file.read(), features='html.parser')
elements = soup.select('li')
params = {}
car_id = 0
for i in range(0, len(elements)):
    element = str(elements[i].getText()).strip()
    key = str(element[:int(element.index(':'))]).strip()
    value = str(element[int(element.index(':') + 1):]).strip()
    if key.__contains__('HQEV'):
        index = key.find('HQEV')
        car_id = key[index:index + 7]
    params[key] = value
for key, value in data.items():
    params[key] = value
print('\n***Finished parsing OTA.html, writing data into spreadsheet......\n')

workbook = xlwings.Book(r'C:\Users\jcglq\Desktop\沧州车辆版本检查反馈表.xlsx')
sheet_name = 'Sheet1'
sheet = workbook.sheets[sheet_name]
pprint.pprint(params)
print('\ncar_id', car_id, '\n')
for row in range(1, 33):
    header_cell = sheet.range('A' + str(row))
    if header_cell.value == car_id:
        for column in range(ord('B'), ord('V')):
            current_cell = sheet.range(chr(column) + str(row))
            standard_cell = sheet.range(chr(column) + str(32))
            param_key = sheet.range(str(chr(column)) + str(1)).value
            current_cell.api.Font.Color = 0
            current_cell.api.Font.Bold = False
            if params.__contains__(param_key):
                current_cell.value = params[param_key]
                if current_cell.value != standard_cell.value:
                    current_cell.api.Font.Color \
                        = 0x0000ff
            elif param_key == 'iso版本':
                current_cell.value = params['pioneer']
                if current_cell.value != sheet.range('B' + str(32)).value:
                    print(current_cell.value, '***', sheet.range('B' + str(32)), '***Color...\n\n')
                    current_cell.api.Font.Color = 0x0000ff
workbook.save()
workbook.close()