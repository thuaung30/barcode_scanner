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
import cv2
import imutils
from imutils.video import VideoStream 


song = AudioSegment.from_mp3('i-demand-attention.mp3')
flag = False
state = True

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
for i in range(0,5):
    timer_dict[timer[i]] = dateTime[i]


# for k,v in timer_dict.items():
#     print(k, v)


# def flush(col):
#     for i in range(2,10):
#         readWorkSheet['{}{}'.format(col, i)] = None


scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('barcodeID-spreadsheet.json', scope)

def detect(image, threshold, morpth_matrix_size, blur_matrix):
    gray = cv2.cvtColor(image,cv2.COLOR_BGR2GRAY)

    ddepth = cv2.cv.CV_32F if imutils.is_cv2() else cv2.CV_32F
    gradX = cv2.Sobel(gray, ddepth = ddepth, dx=1, dy=0, ksize=-1)
    gradY = cv2.Sobel(gray, ddepth = ddepth, dx=0, dy=1, ksize=-1)
    gradient = cv2.subtract(gradX, gradY)
    gradient = cv2.convertScaleAbs(gradient)
    
    
    blurred = cv2.blur(gradient, blur_matrix)
    (_, thresh) = cv2.threshold(blurred, threshold ,255, cv2.THRESH_BINARY)
    

    rect_kernal = cv2.getStructuringElement(cv2.MORPH_RECT, morpth_matrix_size)
   

    rect_image = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, rect_kernal)

    rect_image = cv2.morphologyEx(rect_image, cv2.MORPH_OPEN, rect_kernal)

    cnts = cv2.findContours(rect_image.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cnts = imutils.grab_contours(cnts)
    if len(cnts) == 0:
        return None
    
    c = sorted(cnts, key = cv2.contourArea)[-1]
    x,y,w,h = cv2.boundingRect(c)
#     rect = cv2.minAreaRect(c)
#     box = cv2.cv.BoxPoints(rect) if imutils.is_cv2() else cv2.boxPoints(rect)
#     box = np.int0(box)
    return x,y,w,h

print("[INFO] Video stream starting soon")
vs = VideoStream(src=1).start()
time.sleep(2)
######################################Tuning Parameters############################################
threshold = 160
morph_matrix_size = (14,14)
blur_matrix = (9,9)

code_num = 1

csv = open("barcodes.csv", "w")
found = set()

# font = cv2.FONT_HERSHEY_PLAIN


strippedData = " "

while True:
    try:
        frame = vs.read()
        frame2 = vs.read()
        if frame is None:
            break
        gray = cv2.cvtColor(frame2,cv2.COLOR_BGR2GRAY)

        ddepth = cv2.cv.CV_32F if imutils.is_cv2() else cv2.CV_32F
        gradX = cv2.Sobel(gray, ddepth = ddepth, dx=1, dy=0, ksize=-1)
        gradY = cv2.Sobel(gray, ddepth = ddepth, dx=0, dy=1, ksize=-1)
        gradient = cv2.subtract(gradX, gradY)
        gradient = cv2.convertScaleAbs(gradient)
        
        blurred = cv2.blur(gradient, blur_matrix)
        (_, thresh) = cv2.threshold(blurred, threshold ,255, cv2.THRESH_BINARY)
        

        rect_kernal = cv2.getStructuringElement(cv2.MORPH_RECT, morph_matrix_size)
    

        rect_image = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, rect_kernal)

        rect_image = cv2.morphologyEx(rect_image, cv2.MORPH_OPEN, rect_kernal)
        
        try:
            x,y,w,h = detect(frame, threshold, morph_matrix_size, blur_matrix)
            try:
                bounded_frame = cv2.rectangle(frame.copy(),(x,y),(x+w, y+h), (255,0,0), 2)
            except Exception as e:
                bounded_frame = cv2.rectangle(frame.copy(),(0,0),(10, 10), (255,0,0), 2)
            cropped_barcode = frame[y:y+h,x:x+w]
            cv2.imshow("Cropped", cropped_barcode)
            decodedObjects = pyzbar.decode(cropped_barcode)
            for obj in decodedObjects:
                strippedData = str(obj.data).strip("b''")
                print(strippedData)
                #cv2.putText(cropped_barcode, strippedData, (50, 50), cv2.FONT_HERSHEY_PLAIN, 2, (255, 0, 0), 3)
                if strippedData != " ": 
                    for cell in readWorkSheet[1]:
                        if strippedData == cell.value:
                            column = cell.column
                            columnLetter = chr(64+column)
                            count = readWorkSheet['{}10'.format(columnLetter)].value
                            a = time.time()
                            
                            if a > int(timer_dict[strippedData])+10 and int(count) < 8: 
                                timer_dict[strippedData] = time.time()
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
            cv2.imshow("Capturing barcodes",cropped_barcode)
            key = cv2.waitKey(1)

            if key == ord("q"):
                break
        except Exception as e:
            bounded_frame = frame

    except(KeyboardInterrupt):
        print("Turning off camera.")
        vs.release()
        print("Camera off.")
        print("Program ended.")
        cv2.destroyAllWindows()
        break

print("[INFO] Cleaning up...")
csv.close()
vs.stop()
