import bs4,xlrd,xlwt

# ota_url = 'http://192.168.10.6:8357'
# webbrowser.open(ota_url)
# response = requests.get(ota_url)
# file = open('OTA.html', 'wb')
# for chunk in response.iter_content(100000):
#     file.write(chunk)
# print('\n')
# print(response.text)

file = open('OTA.html')
soup = bs4.BeautifulSoup(file.read())
elements = soup.select('li')
params = {}
for i in range(0, len(elements)):
    element = str(elements[i].getText()).strip()
    key = str(element[:int(element.index(':'))]).strip()
    value = str(element[int(element.index(':') + 1):]).strip()
    params[key] = value
    print(key, value)

print('Finished parsing OTA.html.....\nWriting data into spreadsheet...\n')

workbook = xlrd.open_workbook('版本检查反馈表.xlsx')
