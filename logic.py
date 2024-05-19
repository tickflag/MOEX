from table import XLSXTable
from web import MOEXRequest

class MOEXProgramm:
    def __init__(self, name: str, startDate: str, endDate: str, startIndex=0, pageSize=100):
        self.name = name
        self.startDate = startDate
        self.endDate = endDate
        self.currentDate = startDate
        self.startIndex = startIndex
        self.pageSize = pageSize

        self.limit = None
        self.table = None
        self.resp = None
        self.error = None

        self.url = 'https://iss.moex.com/iss/history/engines/stock/markets/shares/boardgroups/57/securities.json?iss.meta=off&iss.json=extended&callback=JSON_CALLBACK&lang=ru&security_collection=3&date={currentDate}&start={startIndex}&limit={limit}&sort_column=VALUE&sort_order=desc'

    def createTable(self):
        self.table = XLSXTable(f'./{self.name}.xlsx')
        self.table.createTable()
        self.table.loadTable()
    
    def loadTable(self):
        self.table.loadTable()

    def closeTable(self):
        self.table.closeTable()

    def makeNewSheet(self):
        self.table.createNewSheet(self.currentDate)
        self.table.setCurrentSheet(self.currentDate)

    def makeRequest(self, currentDate, startIndex, limit):
        self.resp = MOEXRequest(self.url.format(currentDate=currentDate, startIndex=startIndex, limit=limit))
        self.resp.makeGetRequest()
        self.resp.responseToJson()
        self.setLimit(self.resp.getHistoryCursor()['TOTAL'])

    def testRequest(self):
        self.makeRequest(self.currentDate, 0, 1)

    #set and get for self.*
    def setError(self, error):
        self.error = error
    
    def getError(self):
        return self.error

    def getName(self):
        return self.name

    def setName(self, name):
        self.name = name

    def getStartDate(self):
        return self.startDate

    def setStartDate(self, startDate):
        self.startDate = startDate

    def getEndDate(self):
        return self.endDate
    
    def setEndDate(self, endDate):
        self.endDate = endDate

    def getCurrentDate(self):
        return self.currentDate
    
    def setCurrentDate(self, currentDate):
        self.currentDate = currentDate

    def getStartIndex(self):
        return self.startIndex

    def setStartIndex(self, startIndex):
        self.startIndex = startIndex

    def getPageSize(self):
        return self.pageSize

    def setPageSize(self, pageSize):
        self.pageSize = pageSize

    def getLimit(self):
        return self.limit

    def setLimit(self, limit):
        self.limit = limit

    def getTable(self):
        return self.table

    def getResponse(self):
        return self.response

    
