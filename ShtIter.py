import xlwings as xw


class ShtIter:
    def __init__(self, sht, x1=1, x2=15, yCache=20):
        self.sht = sht
        self.x1 = x1
        self.x2 = x2
        self.yCache = yCache

        self.cache = None
        self.cacheCnt = yCache

        self.totalCnt = 1

    def __iter__(self):
        self.__next__()
        self.__next__()
        return self

    def __next__(self):
        if not (self.cacheCnt < self.yCache):
            self.refillCache()
        self.cacheCnt += 1
        return self.cache.__next__()
        # raise StopIteration

    def refillCache(self):
        self.cacheCnt = 0
        cache = self.sht.range((self.totalCnt, self.x1), (self.totalCnt + self.yCache, self.x2)).value
        self.totalCnt += self.yCache
        self.cache = iter(cache)


def test():
    book = xw.Book('iterTest.xlsx')
    sht = book.sheets.active
    for i, l in enumerate(ShtIter(sht)):
        print(l)
        if i > 100:
            break


if __name__ == '__main__':
    test()
