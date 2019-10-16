import cv2
import numpy as np
import pyzbar.pyzbar as pyzbar
from openpyxl import Workbook
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import urllib.request as req
from urllib.error import URLError
import time
from pydub import AudioSegment
from pydub.playback import play

song = AudioSegment.from_mp3('i-demand-attention.mp3')

workbook = Workbook()
ws = workbook.active

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('barcodeID-spreadsheet.json', scope)


cap = cv2.VideoCapture(1)
font = cv2.FONT_HERSHEY_PLAIN

strippedData = " "

def interneton():
    try:
        req.urlopen('https://www.google.com', timeout=1)
        return True
    except URLError:
        return False
        
def updategspreadsheet():
    _, frame = cap.read()
    fframe = cv2.flip(frame, 1)
    decodedObjects = pyzbar.decode(frame)
    for obj in decodedObjects:
        strippedData = str(obj.data).strip("b''")
        cv2.putText(fframe, strippedData, (50, 50), font, 2, (255, 0, 0), 3)
    cv2.imshow("Frame", fframe)
    return strippedData

if interneton() == True:
    client = gspread.authorize(creds)
    sheet = client.open('barcodeID spreadsheet').sheet1
    while True:
        _, frame = cap.read()
        fframe = cv2.flip(frame, 1)
        decodedObjects = pyzbar.decode(frame)
        for obj in decodedObjects:
            strippedData = str(obj.data).strip("b''")
            cv2.putText(fframe, strippedData, (50, 50), font, 2, (255, 0, 0), 3)
        cv2.imshow("Frame", fframe)
        if strippedData != " ": 
            sheet.update_cell( 1, 1, strippedData)
            ws['A1'] = strippedData
            workbook.save("sample.xlsx")
            play(song)
            time.sleep(5)
            strippedData = " "
        key = cv2.waitKey(1)
        if key == 27:
            break

else: 
    while True: 
        _, frame = cap.read()
        fframe = cv2.flip(frame, 1)
        decodedObjects = pyzbar.decode(frame)
        for obj in decodedObjects:
            strippedData = str(obj.data).strip("b''")
            cv2.putText(fframe, strippedData, (50, 50), font, 2, (255, 0, 0), 3)
        cv2.imshow("Frame", fframe)
        if strippedData != " ":
            ws['A1'] = strippedData
            workbook.save("sample.xlsx")
            play(song)
            time.sleep(5)
            strippedData = " "
        key = cv2.waitKey(1)
        if key == 27:
            break


