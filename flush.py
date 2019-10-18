from openpyxl import Workbook
from openpyxl import load_workbook
import gspread
from oauth2client.service_account import ServiceAccountCredentials


readWorkBook = load_workbook('attendanceCheck.xlsx')
readWorkSheet = readWorkBook.worksheets[0]


scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('barcodeID-spreadsheet.json', scope)
client = gspread.authorize(creds)
sheet = client.open('barcodeID spreadsheet').sheet1

def coder(a):
    if a <= 9:
        b = 7-a
        if b < 0:
            c = 10 + b 
            return '119000{}42198{}'.format(a, c)
        return '119000{}42198{}'.format(a, b)
    else: 
        b = 14 - a
        if b < 0:
            c = 10 + b
            return '11900{}42198{}'.format(a, c)
        return '11900{}42198{}'.format(a, b)


def flush():
    identity = int(input('Enter student ID.\n'))
    barCode = coder(identity)
    for cell in readWorkSheet[1]:
        if cell.value == barCode:
            column = cell.column
            columnLetter = chr(64+column) 
            for i in range(2,10):
                readWorkSheet['{}{}'.format(columnLetter, i)] = ' '
                sheet.update_cell(i, column, ' ')
            readWorkSheet['{}10'.format(columnLetter)] = '0'
            sheet.update_cell(10, column, '0')
            print('Flushed.')
    readWorkBook.save("attendanceCheck.xlsx")
        

flush()