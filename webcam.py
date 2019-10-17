import cv2
import numpy as np
import pyzbar.pyzbar as pyzbar
from openpyxl import Workbook
from openpyxl import load_workbook
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import httplib2
import time 
import datetime
from pydub import AudioSegment
from pydub.playback import play


song = AudioSegment.from_mp3('i-demand-attention.mp3')
flag = False


# workbook = Workbook()
# ws = workbook.active
readWorkBook = load_workbook('attendanceCheck.xlsx')
readWorkSheet = readWorkBook.worksheets[0]

timer = []
dateTime = []
for cell in readWorkSheet[1]:
    if cell.value != 'BarCodeID':
        timer.append(cell.value)
for cell in readWorkSheet[11]:
    if cell.value != 'DateTime':
        dateTime.append(datetime.datetime.strptime(str(cell.value), '%a %b %d %H:%M:%S %Y').timestamp())
timer_dict = {}
for i in range(0,9):
    timer_dict[timer[i]] = dateTime[i]

# for k,v in timer_dict.items():
#     print(k, v)


scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('barcodeID-spreadsheet.json', scope)


cap = cv2.VideoCapture(1)
font = cv2.FONT_HERSHEY_PLAIN


strippedData = " "
        

while True:
    _, frame = cap.read()
    fframe = cv2.flip(frame, 1)
    decodedObjects = pyzbar.decode(frame)
    for obj in decodedObjects:
        strippedData = str(obj.data).strip("b''")
        #cv2.putText(fframe, strippedData, (50, 50), font, 2, (255, 0, 0), 3)
    cv2.imshow("Frame", fframe)
    if strippedData != " ": 
        for cell in readWorkSheet[1]:
            if strippedData == cell.value:
                a = time.time()
                if a > int(timer_dict[strippedData])+30: 
                    timer_dict[strippedData] = time.time()
                    column = cell.column
                    columnLetter = chr(64+column)
                    count = readWorkSheet['{}10'.format(columnLetter)].value
                    count1 = str(int(count)+1)
                    readWorkSheet['{}10'.format(columnLetter)] = count1
                    readWorkSheet['{}{}'.format(columnLetter, int(count1)+1)] = 'Present'
                    readWorkSheet['{}11'.format(columnLetter)] = time.ctime(timer_dict[strippedData])
                    play(song)
                    flag = True
        if flag == True:
            try:
                client = gspread.authorize(creds)
                sheet = client.open('barcodeID spreadsheet').sheet1
                sheet.update_cell(int(count1)+1, column, 'Present')
                sheet.update_cell(10, column, count1)
                sheet.update_cell(11, column, time.ctime(timer_dict[strippedData]))
            except (httplib2.ServerNotFoundError, ConnectionError):
                pass
        flag = False
        readWorkBook.save("attendanceCheck.xlsx")
        strippedData = " "
    key = cv2.waitKey(1)
    if key == 27:
        break