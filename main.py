from table import XLSXTable
from web import Request, MOEXRequest
from logic import MOEXProgramm

from datetime import date, timedelta
import datetime

class ImplementDate:
    def __init__(self, date: str):
        self.date = date
        self.year = None
        self.month = None
        self.day = None

    def parseDate(self):
        date = self.date.split('-')
        self.year = int(date[0])
        self.month = int(date[1].lstrip('0'))
        self.day = int(date[2].lstrip('0'))

    def getYear(self):
        return self.year
    
    def getMonth(self):
        return self.month
    
    def getDay(self):
        return self.day

def date_range_list(start_date, end_date):
    date_list = []
    curr_date = start_date
    while curr_date <= end_date:
        date_list.append(curr_date.__str__())
        curr_date += timedelta(days=1)
    return date_list

def solution(name, startDate, endDate, startIndex, pageSize):

    date_list = date_range_list(startDate, endDate)
    moex = MOEXProgramm(name, startDate, endDate, startIndex, pageSize)
    moex.createTable()
    print(f"INFO: \nSTARTDATE: {moex.getStartDate()} \nENDDATE: {moex.getEndDate()} \nELEMENTS = {len(date_list)} \n{'-' * 30}")
    for item in date_list:
        moex.setCurrentDate(item)
        moex.testRequest()
        moex.table.setLastIndex(1)
        if (moex.resp.getStatusCode() != 200):
            print(moex.resp.getResponseCode())
            print('!'*30)
            print("ERROR DURING REQUEST!\n TRY LATER!")
            print(f"ERROR APPIARED ON DATE {moex.getCurrentDate()}")
            print('!'*30)
            moex.setError("UNABLE TO GET RESPONSE FROM MOEX")
            break

        if (moex.getLimit() != 0):
            moex.makeNewSheet()
            moex.table.addSheet()
            while(moex.getStartIndex() <= moex.getLimit()):
                moex.makeRequest(moex.getCurrentDate(), moex.getStartIndex(), moex.getLimit())
                data = moex.resp.getHistory()
                moex.table.setValueOnSheet(f'A1', 'SECID')
                moex.table.setColumnWidth('A', 10)
                moex.table.setValueOnSheet(f'B1', 'SHORTNAME')
                moex.table.setColumnWidth('B', 15)
                moex.table.setValueOnSheet(f'C1', 'NUMTRADES')
                moex.table.setColumnWidth('C', 15)
                moex.table.setValueOnSheet(f'D1', 'VALUE')
                moex.table.setColumnWidth('D', 30)
                moex.table.setValueOnSheet(f'E1', 'WAPRICE')
                moex.table.setColumnWidth('E', 15)
                moex.table.setValueOnSheet(f'F1', 'OPEN')
                moex.table.setColumnWidth('F', 15)
                moex.table.setValueOnSheet(f'G1', 'LOW')
                moex.table.setColumnWidth('G', 15)
                moex.table.setValueOnSheet(f'H1', 'HIGH')
                moex.table.setColumnWidth('H', 15)
                moex.table.setValueOnSheet(f'I1', 'CLOSE')
                moex.table.setColumnWidth('I', 15)
    
                for element in data:
                    index = moex.table.getLastIndex() + 1
                    moex.table.setValueOnSheet(f'A{index}', element['SECID'])
                    moex.table.setValueOnSheet(f'B{index}', element['SHORTNAME'])
                    moex.table.setValueOnSheet(f'C{index}', element['NUMTRADES'])
                    moex.table.setValueOnSheet(f'D{index}', element['VALUE'])
                    moex.table.setValueOnSheet(f'E{index}', element['WAPRICE'])
                    moex.table.setValueOnSheet(f'F{index}', element['OPEN'])
                    moex.table.setValueOnSheet(f'G{index}', element['LOW'])
                    moex.table.setValueOnSheet(f'H{index}', element['HIGH'])
                    moex.table.setValueOnSheet(f'I{index}', element['CLOSE'])
                    moex.table.setLastIndex(index)
            
                moex.setStartIndex(moex.getStartIndex() + moex.getPageSize())
            print(f"{moex.getCurrentDate()} COMPLITED SUCCESFUL, got {moex.getLimit()} elements")
        else:
            print(f"{moex.getCurrentDate()} SKIPPED")
        moex.setStartIndex(0)
        moex.table.saveTable()
    return moex, len(date_list)

def chooseCustomRange():
    print("DATE FORMAT: YYYY.MM.DD")
    print("START DATE: ", end='')
    start_date = input()
    print("END DATE: ", end='')
    end_date = input()
    print('-' * 30)
    name = f"[{start_date}]-[{end_date}]"
    startDate = ImplementDate(start_date)
    startDate.parseDate()
    startDate = datetime.date(year=startDate.getYear(), month=startDate.getMonth(), day=startDate.getDay())
    endDate = ImplementDate(end_date)
    endDate.parseDate()
    endDate = datetime.date(year=endDate.getYear(), month=endDate.getMonth(), day=endDate.getDay())
    main(name, startDate, endDate)

def previousDay():
    name = f"[{datetime.date.today()}]"
    startDate = datetime.date.today() - timedelta(days=1)
    endDate = startDate
    main(name, startDate, endDate)

def previousWeek():
    startDate = datetime.date.today() - timedelta(weeks=1)
    endDate = datetime.date.today() - timedelta(days=1)
    name = f"[{startDate}]-[{endDate}]"
    main(name, startDate, endDate)

def previousMonth():
    startDate = datetime.date.today() - timedelta(weeks=4)
    endDate = datetime.date.today() - timedelta(days=1)
    name = f"[{startDate}]-[{endDate}]"
    main(name, startDate, endDate)

def main(name: str, startDate, endDate):
    pageSize = 100
    startIndex = 0
    moex, l = solution(name, startDate, endDate, startIndex, pageSize)
    print('-'*30)
    print(f"INFO: \nSTATUS: COMPLITED \nERRORS: {moex.getError()} \nSAVED TO: {moex.getName()}.xlsx \nSHEETS: {moex.table.getSheets()} \nSKIPPED DATES: {l - moex.table.getSheets()}")

def chooseRange():
    print("CHOOSE MODE")
    print("1. PREVIOUS DAY \n2. PREVIOUS WEEK \n3. PREVIOUS MONTH \n4. CUSTOM RANGE")
    while (True):
        choice = int(input("INPUT: "))
        if (choice == 1):
            previousDay()
            break
        elif (choice == 2):
            previousWeek()
            break
        elif (choice == 3): 
            previousMonth()
            break
        elif (choice == 4):
            chooseCustomRange()
            break
        else:
            print("PLEASE CHOOSE NUMBER FROM 1 TO 4")



if __name__ == '__main__':
    chooseRange()
