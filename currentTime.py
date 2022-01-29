'''currentTime.py - get the current time and return in a string form.

Typical usage example:

currentTime.getTime()

'''

import time


def getTime():
    cT = time.strftime("%Y-%m-%d %H.%M", time.localtime())
    return cT

if __name__ == '__main__':
    print(getTime())
    sieve = [True] * 100
    for i in range(2, 100):
        if sieve[i]:
            print(i)
            for j in range(i*i, 100, i):
                sieve[j] = False
