import json
import os
import xlwings as xw

from ShtIter import ShtIter
from SpTools import getIdentifiers

IDENTIFIER = getIdentifiers()
END_WORDS = ['#끝']
DEFAULT_IDX = '0'


class SpOut:
    def __init__(self, sht):
        self.nOrders = []
        self.nCustomer = None
        self.sht = sht
        self.parsedData = {}

    def parseSht(self):
        for line in self:
            if line[2] in IDENTIFIER:
                self.nOrders = []
                self.parsedData[line[2]] = self.nOrders
                continue
            if line[2] in END_WORDS:
                return self.parsedData
            if (line[6] is not None) and (line[7] is not None):
                self.nCustomer = self.parseCustomer(line)
            if ((line[3] is None) and (line[4] is None) and (line[10] is None)) or (line[11] == 0) or (line[11] == None):
                continue
            order = {'goods': line[3], 'amount': line[4], 'customer': self.nCustomer, 'perPrice': line[10],
                     'totalPrice': line[11]}
            self.nOrders.append(order)

        return self.parsedData

    def parseCustomer(self, line):
        for i in [6, 7]:
            if '#' in line[i]:
                idx = line[i].split('#')[-1]
                break
        else:
            idx = DEFAULT_IDX
        return {'branch': line[6], 'name': line[7], 'addr': line[8], 'contact': line[9], 'idx': idx}

    def __iter__(self):
        return ShtIter(self.sht)

    @classmethod
    def getFilename(cls, names):
        fEx = '.xlsx'
        include = '출고'
        for f in names:
            if f.endswith(fEx) and (include in f) and (not f.startswith('~')):
                return f

    @classmethod
    def getBook(cls, fName):
        return xw.Book(fName)

    @classmethod
    def getSht(cls, book, shtName=None):
        if shtName is None:
            return book.sheets.active
        else:
            return book.sheets[shtName]

    @classmethod
    def getShtByFnames(cls, fnames):
        name = cls.getFilename(fnames)
        book = cls.getBook(name)
        sht = cls.getSht(book)
        return sht


def test():
    sht = SpOut.getShtByFnames(os.listdir())
    spOut = SpOut(sht)
    parsedData = spOut.parseSht()
    print(json.dumps(parsedData, ensure_ascii=False))


if __name__ == '__main__':
    test()
