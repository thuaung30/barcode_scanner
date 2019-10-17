from openpyxl import Workbook
wb = Workbook()


# grab the active worksheet
ws = wb.active

ws.append(['Name', 'barCodeID', 'level', 'phone'])


class student():
    def __init__(self, name, barCodeID, level, phone):
        self.name = name
        self.barCodeID = barCodeID
        self.level = level
        self.phone = phone

    def addValue(self):
        ws.append([self.name, self.barCodeID, self.level, self.phone])


s1 =student('Aung Kyaw Hein', '0001', 'K1', '0979798834')
s1.addValue()


s2 = student('Aung Htet', '0002', 'K2', '0979798835')
s2.addValue()


s3 = student('Zaw Nay Lin', '0003', 'K2', '0979798836')
s3.addValue()


# Save the file
wb.save("studentData.xlsx")