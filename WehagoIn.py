import os

import xlwings as xw
import time

from SpOut import SpOut


class WehagoIn:
    def __init__(self):
        self.book = xw.Book()
        self.sht = self.book.sheets.active
        self.setTitle()
        self.setDate()
        self.fromData = []

    def setTitle(self):
        titles = ['년도', '월', '일', '매입매출구분(1 - 매출 / 2 - 매입)', '과세유형', '불공제사유', '신용카드거래처코드', '신용카드사명', '신용카드(가맹점)번호',
                  '거래처명', '사업자(주민)등록번호', '공급가액', '부가세', '품명', '전자세금(1.전자)', '기본계정', '상대계정', '현금영수증 승인번호']
        self.sht.name = '매출자료 & 매입자료'
        self.sht.range(1, 1).value = titles

    def defaultLine(self):
        res = [None] * 18
        res[0], res[1], res[2], res[3], res[4] = self.year, self.month, self.date, '1', '13'
        res[15], res[16] = '401', '108'
        return res

    def fromSpOut(self, parsedData):
        for k, orders in parsedData.items():
            for order in orders:
                outOrder = self.defaultLine()
                outOrder[9] = order['customer']['branch']
                outOrder[10] = order['customer']['idx'].zfill(6) + '-0000000'
                outOrder[11] = order['totalPrice']
                outOrder[13] = order['goods']
                self.fromData.append(outOrder)

        self.sht.range(2, 1).value = self.fromData

    def setDate(self):
        t = time.localtime(time.time())
        self.year, self.month, self.date = time.strftime('%Y', t), time.strftime('%m', t), time.strftime('%d', t)


def test():
    sht = SpOut.getShtByFnames(os.listdir())
    spOut = SpOut(sht)
    parsedData = spOut.parseSht()
    wehagoIn = WehagoIn()
    wehagoIn.fromSpOut(parsedData)


if __name__ == '__main__':
    test()
