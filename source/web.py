import requests


class Request:
    def __init__(self, url: str):
        self.url = url
        self.response = None
        self.statusCode = None

    def responseToJson(self):
        self.response = self.response.json()
    
    def makeGetRequest(self):
        self.response = requests.get(self.url)
        self.statusCode = self.response.status_code

    #get and set self.*
    def setUrl(self, url):
        self.url = url

    def getUrl(self):
        return self.url

    def getStatusCode(self):
        return self.statusCode
    
    def getResponse(self):
        return self.response


class MOEXRequest(Request):
    def getHistoryCursor(self):
        return self.response[1]['history.cursor'][0]

    def getHistory(self):
        return self.response[1]['history']

