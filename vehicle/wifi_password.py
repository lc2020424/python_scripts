import xlwings

workbook = xlwings.Book(r'C:\Users\jcglq\Desktop\wifi-password.xlsx')
sheet = workbook.sheets['Sheet1']


def show_password(num: int) -> str:
    for row in range(3, 33):
        cell = sheet.range('A' + str(row))
        if cell.value.find(str(num)) >= 0:
            print(sheet.range('B' + str(row)).value, sheet.range('C' + str(row)).value)
            break


show_password(554)
workbook.close()