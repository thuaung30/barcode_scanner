import cv2
import numpy as np
import pyzbar.pyzbar as pyzbar
from openpyxl import Workbook
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# workbook = Workbook()
# ws = workbook.active

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('barcodeID-spreadsheet.json', scope)
client = gspread.authorize(creds)
sheet = client.open('barcodeID spreadsheet').sheet1

cap = cv2.VideoCapture(1)
font = cv2.FONT_HERSHEY_PLAIN


while True:
    _, frame = cap.read()
    fframe = cv2.flip(frame, 1)
    decodedObjects = pyzbar.decode(frame)
    for obj in decodedObjects:
        # print('Type : ', obj.type)
        # print("Data = ", str(obj.data).strip("b''"))
        strippedData = str(obj.data).strip("b''")
        sheet.update_cell(1, 1, strippedData)
        cv2.putText(fframe, strippedData, (50, 50), font, 2, (255, 0, 0), 3)

        # ws['A1'] = strippedData
        # workbook.save("sample.xlsx")
        # else:
        #     pass     

    cv2.imshow("Frame", fframe)

    key = cv2.waitKey(1)
    if key == 27:
        break