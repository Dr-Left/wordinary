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
