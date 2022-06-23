def getIdentifiers():
    return sorted(_getIdentifiers())


def _getIdentifiers():
    suffix = '#'
    ss1 = ['월', '화', '수', '목', '금', '토', '일']
    ss2 = ['요일', '욜']
    ss3 = ['택배', '배송']
    etc = ['행복앤미소', '수건어물']
    for s1 in ss1:
        for s2 in ss2:
            for s3 in ss3:
                yield '{}{}{}{}'.format(suffix, s1, s2, s3)

    for s in etc:
        yield '{}{}'.format(suffix, s)


def test():
    s = sorted(getIdentifiers())
    print(type(s))
    print(s)


if __name__ == '__main__':
    test()
