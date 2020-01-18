import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('barcodeID-spreadsheet.json', scope)

client = gspread.authorize(creds)
sheet = client.open('barcodeID spreadsheet').sheet1
x = 7

sheet.update_cell( 1, 1, 'BarCodeID')
sheet.update_cell(26, 1, 'Count')
sheet.update_cell(27, 1, 'DateTime')
for column in range(2,7):
    sheet.update_cell(1, column, ' ')
    sheet.update_cell(26, column, '0')
    sheet.update_cell(27, column, '0')

for column in range(2,7):
    sheet.update_cell( 1, column, '191100{}'.format(column-1))

for row in range(2,26):
    sheet.update_cell(row, 1, 'Day{}'.format(row-1))