import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('barcodeID-spreadsheet.json', scope)

client = gspread.authorize(creds)
sheet = client.open('barcodeID spreadsheet').sheet1
x = 7

sheet.update_cell( 1, 1, 'BarCodeID')
sheet.update_cell(10, 1, 'Count')
sheet.update_cell(11, 1, 'DateTime')
for column in range(2,11):
    sheet.update_cell(10, column, '0')
for column in range(2,11):
    x1 = x-(column-1)
    if x1 < 0:
        x1 = 10 + x1
    sheet.update_cell( 1, column, '119000{}42198{}'.format(column-1, x1))

for row in range(2,10):
    sheet.update_cell(row, 1, 'Day{}'.format(row-1))